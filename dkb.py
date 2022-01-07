#!/usr/bin/python3
# -*- coding: utf-8 -*-
# DKB Credit card transaction QIF exporter
# Copyright (C) 2013 Christian Hoffmann <mail@hoffmann-christian.info>
#
# Inspired by Jens Herrmann <jens.herrmann@qoli.de>,
# but written using modern tools (argparse, csv reader, mechanize,
# BeautifulSoup)
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU Affero General Public License as
# published by the Free Software Foundation, either version 3 of the
# License, or (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU Affero General Public License for more details.
#
# You should have received a copy of the GNU Affero General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.

import re
import os
import csv
import sys
import pickle
import logging
import mechanize
import time
import unittest
from http.cookiejar import MozillaCookieJar


class RecordingBrowser(mechanize.Browser):
    _recording_path = None
    _recording_enabled = False
    _playback_enabled = False
    _intercept_count = 0

    def enable_recording(self, path):
        self._recording_path = path
        self._recording_enabled = True

    def enable_playback(self, path):
        self._recording_path = path
        self._playback_enabled = True

    def open(self, *args, **kwargs):
        try:
            return self._intercept_call('open', *args, **kwargs)
        except Exception as e:
            print(e)
            raise e

    def _intercept_call(self, method, *args, **kwargs):
        if self._playback_enabled:
            self._intercept_count += 1
            return self._read_recording()

        func = getattr(mechanize.Browser, method)
        ret = func(self, *args, **kwargs)
        if self._recording_enabled:
            self._do_record()
        return ret

    def _do_record(self):
        """
        Writes the current HTML to disk if dumping is enabled.
        Useful for offline testing.
        """
        data = {}
        resp = self.response()
        if not resp:
            return
        data['data'] = resp.get_data()
        data['code'] = resp.code
        data['msg'] = resp.msg
        data['headers'] = resp.info().items()
        data['url'] = resp.geturl()

        self._intercept_count += 1
        dump_path = '%s/%d.pickle' % (self._recording_path, self._intercept_count)
        with open(dump_path, 'wb') as f:
            pickle.dump(data, f)

    def _read_recording(self):
        dump_path = '%s/%d.pickle' % (self._recording_path, self._intercept_count)
        if not os.path.exists(dump_path):
            self._intercept_count += 1
            dump_path = '%s/%d.pickle' % (self._recording_path, self._intercept_count)
        with open(dump_path, 'rb') as f:
            data = pickle.load(f)
            if not data:
                return self.set_response(None)
            resp = mechanize.make_response(**data)
            self.set_response(resp)
            return resp


logger = logging.getLogger(__name__)


class DkbScraper(object):
    BASEURL = "https://www.dkb.de/-"

    def __init__(self, record_html=False, playback_html=False, session_persistence_file=None):
        self.br = RecordingBrowser()
        dump_path = os.path.join(os.path.dirname(__file__), 'dumps')
        if record_html:
            self.br.enable_recording(dump_path)
        if playback_html:
            self.br.enable_playback(dump_path)
        self.session_persistence_file = session_persistence_file

    def login(self, userid, get_pin_callback):
        """
        Create a new session by submitting the login form

        @param str userid
        @param callable get_pin_callback
        """
        if self.try_persisted_session():
            return
        logger.info("Starting login as user %s...", userid)
        br = self.br

        # we are not a spider, so let's ignore robots.txt...
        br.set_handle_robots(False)
        br.addheaders = [('User-Agent','Mozilla/5.0 (X11; Linux x86_64; rv:76.0) Gecko/20100101 Firefox/76.0')]

        # Although we have to handle a meta refresh, we disable it here
        # since mechanize seems to be buggy and will be stuck in a
        # long (infinite?) sleep() call
        br.set_handle_refresh(False)

        br.open(self.BASEURL + '?$javascript=disabled')

        # select login form:
        br.form = list(br.forms())[1]

        pin = get_pin_callback()

        br.set_all_readonly(False)
        br.form["j_username"] = userid
        br.form["j_password"] = pin
        br.form["jsEnabled"] = "false"
        br.form["browserName"] = "Firefox"
        br.form["browserVersion"] = "76.0"
        br.form["screenWidth"] = "1000"
        br.form["screenHeight"] = "800"
        br.form["osName"] = "UNIX"
        response = br.submit()
        if ("Wechseln Sie in die <strong>DKB-Banking-App</strong> und best" in response.read().decode('utf-8') ):
            self.confirm_app_login()
        else:
            self.confirm_tan_login()

    def try_persisted_session(self):
        if not self.session_persistence_file:
            return
        try:
            self.br.set_cookiejar(MozillaCookieJar(self.session_persistence_file))
            self.br.cookiejar.load(ignore_discard=True)
        except Exception as exc:
            logger.debug('Failed to load persistence file %r', self.session_persistence_file, exc_info=exc)
            return
        logger.debug('Loaded cookiejar')
        return self.is_logged_in()

    def confirm_app_login(self):
        logger.info("DKB-Banking-App detected, waiting for in-app login confirmation...")
        br = self.br
        # The following loop executes the verification sequence for about 2 minutes.
        for x in range(30):
            time.sleep(2)
            if x >= 29:
                print("Authentication timed out")
                quit()
            if not ("WAITING" in br.open('https://www.dkb.de/DkbTransactionBanking/content/LoginWithBoundDevice/LoginWithBoundDeviceProcess/confirmLogin.xhtml?$event=pollingVerification').read().decode('utf-8')):
                break
        br.open(self.BASEURL + "?$javascript=disabled")
        br.select_form(name="confirmForm")
        br.submit()

    def confirm_tan_login(self):
        logger.info("Attempting TAN login")
        br = self.br
        br.form = list(br.forms())[2]
        #FIXME we should check which page we are on...
        try:
            form = self._get_tan_input_form()
        except RuntimeError:
            # if we don't find the tan field, we're probably at the empty form
            br.submit()
            form = self._get_tan_input_form()

        br.form = form
        startcode = re.search("Startcode [0-9]{8}", br.response().get_data())
        if startcode: # using chipTAN
            print(startcode.group())
        # else: using dkbapp
        # TODO check for Startcode
        br.form["tan"] = self.ask_for_tan()
        br.submit()

        try:
            # if we find the tan field after submitting, the TAN was wrong
            self._get_tan_input_form()
        except RuntimeError:
            br.open(self.BASEURL + "?$javascript=disabled")
            return
        raise RuntimeError("TAN seems to be wrong")

    def ask_for_tan(self):
        tan = ""
        import os
        if os.isatty(0):
            while not tan.strip():
                tan = input('TAN: ')
        else:
            tan = sys.stdin.read().strip()
        return tan

    def _get_tan_input_form(self):
        """
        Internal.

        Returns the tan input form object (mechanize)
        """
        for form in self.br.forms():
            try:
                form.find_control(name="tan")
                return form
            except Exception:
                continue

        raise RuntimeError("Unable to find tan input form")

    def is_logged_in(self):
        self.br.open(self.BASEURL)
        try:
            link = self.br.find_link(text='Finanzstatus')
            logger.debug('Session still valid')
            return True
        except mechanize.LinkNotFoundError:
            logger.debug('Session invalid, will re-login')
            return False

    def credit_card_transactions_overview(self):
        """
        Navigates the internal browser state to the credit card
        transaction overview menu
        """
        logger.info("Navigating to 'Umsätze'...")
        try:
            return self.br.follow_link(url_regex='banking/finanzstatus/kontoumsaetze')
        except Exception:
            raise RuntimeError('Unable to find link Umsätze -- '
                               'Maybe the login went wrong?')

    def _get_transaction_selection_form(self):
        """
        Internal.

        Returns the transaction selection form object (mechanize)
        """
        for form in self.br.forms():
            try:
                form.find_control(name="slAllAccounts", type='select')
                return form
            except Exception:
                continue

        raise RuntimeError("Unable to find transaction selection form")

    def _select_all_transactions_from(self, form, from_date, to_date):
        """
        Internal.

        Checks the radio box "Alle Umsätze vom" and populates the
        "from" and "to" with the given values.

        @param mechanize.HTMLForm form
        @param str from_date dd.mm.YYYY
        @param str to_date dd.mm.YYYY
        """
        try:
            radio_ctrl = form.find_control("filterType")
        except Exception:
            raise RuntimeError("Unable to find search period radio box")

        form[radio_ctrl.name] = [u'DATE_RANGE']

        try:
            from_item = form.find_control(label="vom")
        except Exception:
            raise RuntimeError("Unable to find 'vom' date field")

        from_item.value = from_date

        try:
            to_item = form.find_control(label="bis")
        except Exception:
            raise RuntimeError("Unable to find 'to' date field")

        to_item.value = to_date

    def _select_credit_card(self, form, cardid):
        """
        Internal.

        Selects the correct credit card from the dropdown menu in the
        transaction selection form.

        @param mechanize.HTMLForm form
        @param str cardid: last 4 digits of the relevant card number
        """
        try:
            cc_list = form.find_control(name="slAllAccounts", type='select')
        except Exception:
            raise RuntimeError("Unable to find credit card selection form")

        for item in cc_list.get_items():
            # find right credit card...
            for label in item.get_labels():
                pattern = r'\b\S{12}(?<=%s)\b' % re.escape(cardid)
                if re.search(pattern, label.text, re.I):
                    cc_list.value = [item.name]
                    return

        raise RuntimeError("Unable to find the right credit card")

    def select_transactions(self, cardid, from_date, to_date):
        """
        Changes the current view to show all transactions between
        from_date and to_date for the credit card identified by the
        given card id.

        @param str cardid: last 4 digits of your credit card's number
        @param str from_date dd.mm.YYYY
        @param str to_date dd.mm.YYYY
        """
        br = self.br
        logger.info("Selecting transactions in time frame %s - %s...",
                    from_date, to_date)

        br.form = form = self._get_transaction_selection_form()
        self._select_credit_card(form, cardid)
        # we need to reload so that we get the credit card form:
        br.submit()

        br.form = form = self._get_transaction_selection_form()
        self._select_all_transactions_from(form, from_date, to_date)

        # add missing $event control
        br.form.new_control('hidden', '$event', {'value': 'search'})
        br.form.fixup()
        br.submit()

    def get_transaction_csv(self):
        """
        Returns a file-like object which contains the CSV data,
        selected by previous calls.

        @return file-like response
        """
        logger.info("Requesting CSV data...")
        self.br.follow_link(url_regex='csv')
        csv = self.br.response().read()
        #self.br.back()
        return csv

    def close(self):
        """
        Properly closes the session.

        If no persistence is activated, performs a logout.
        If persistence is requested, saves the session cookies and performs NO logout.
        """
        if self.session_persistence_file:
            logger.debug('Writing persistence file')
            try:
                self.br.cookiejar.save(ignore_discard=True)
                return
            except Exception as exc:
                logger.info('failed to write persistence file %r', self.session_persistence_file, exc_info=exc)
                # fall through and perform regular logout:
        logger.debug('Performing logout')
        self.br.open(self.BASEURL + "?$javascript=disabled")
        self.br.follow_link(text='Abmelden')


class DkbConverter(object):
    """
    A DKB transaction CSV to QIF converter

    Tested with GnuCash
    """

    # The financial software's target account (such as Aktiva:VISA)
    DEFAULT_CATEGORY = None

    # QIF-internal card name
    CREDIT_CARD_NAME = 'VISA'

    # input charset
    INPUT_CHARSET = 'latin1'

    # QIF output charset
    OUTPUT_CHARSET = 'utf-8'

    # Length of the pre-amble (non-CSV headers), including the
    # CSV head line
    SKIP_LINES = 8

    # Column Definitions: Which values can be found in which columns?
    COL_DATE = 1
    COL_VALUTA_DATE = 2
    COL_DESC = 3
    COL_VALUE = 4
    COL_INFO = 5

    # Number of fields for the line to be recognized as a valid
    # transaction line:
    REQUIRED_FIELDS = 5

    def __init__(self, csv_text, default_category=None, cc_name=None):
        """
        Constructor

        @param str csv_text
        @param str default_category:
            Category in your financial software
            (an account such as Aktiva:Visa)
        """
        self.csv_text = csv_text.decode(self.INPUT_CHARSET)
        self.DEFAULT_CATEGORY = default_category
        self.CREDIT_CARD_NAME = cc_name or 'VISA'

    def format_date(self, line):
        """
        Extracts the date from the given line and
        converts it from DD.MM.YYYY to MM/DD/YYYY

        @param list line
        @return str
        """
        date_re = re.compile('.*?(\d{1,2})\.(\d{1,2})\.(\d{2,4}).*?')
        if date_re.match(line[self.COL_VALUTA_DATE]):
            # use valuta date if available
            field = self.COL_VALUTA_DATE
        else:
            # ... default to regular date column otherwise:
            field = self.COL_DATE
        return date_re.sub(r'\2/\1/\3', line[field])

    def format_value(self, line):
        """
        Extracts the value (such as 3,00 or -100,00) from the given
        line, removes any dots and replaces the comma by a dot.

        @param list line
        @return str
        """
        return (line[self.COL_VALUE]
                .strip()
                .replace('.', '')  # 1.000 -> 1000
                .replace(',', '.'))  # 0,83 -> 0.83

    def format_description(self, line):
        """
        Extracts the description from the given line and strips
        any whitespace.

        @param list line
        @return str
        """
        return line[self.COL_DESC].strip()

    def format_info(self, line):
        """
        Extracts any additional info (such as different currencies)
        from the given line and strips any whitespace.

        @param list line
        @return str
        """
        return line[self.COL_INFO].strip()

    def get_category(self, line):
        """
        Returns the best-fitting category for the given line.
        Currently, we always return the default category, but this would
        be the place for guessing algorithms.

        @param list line
        @return str
        """
        return self.DEFAULT_CATEGORY

    def get_qif_lines(self):
        """
        Does the actual CSV to QIF conversion and returns an iterator
        over all required QIF lines.
        No line separator is included.

        @return iterator
        """
        logger.info("Running csv->qif conversion...")
        yield '!Account'
        yield 'N' + self.CREDIT_CARD_NAME
        yield '^'
        yield '!Type:Bank'
        lines = self.csv_text.split('\n')
        reader = csv.reader(lines, delimiter=";")
        for x in range(self.SKIP_LINES):
            reader.__next__()
        for line in reader:
            if len(line) < self.REQUIRED_FIELDS:
                continue
            if len(line[self.COL_VALUTA_DATE]) == 0:
                continue
            yield 'D%s' % self.format_date(line)
            yield 'T%s' % self.format_value(line)
            yield 'M%s' % self.format_description(line)
            if line[self.COL_INFO].strip():
                yield 'M%s' % self.format_info(line)
            category = self.get_category(line)
            if category:
                yield 'L%s' % category
            yield '^'

    def export_to(self, path):
        """
        Writes the QIF version of the already stored csv text to the
        given path.

        @param str path
        """
        logger.info("Exporting qif to %s", path)
        with open(path, "wb") as f:
            for line in self.get_qif_lines():
                f.write((line + "\n").encode(self.OUTPUT_CHARSET))


if __name__ == '__main__':
    from getpass import getpass
    from argparse import ArgumentParser
    from datetime import date

    cli = ArgumentParser(description="Download VISA card transactions from DKB account. Specify exactly one userid and same number of cardid, output, qif-account, from-date and to-date arguments (qif-account and to-date are optional).")
    cli.add_argument("--userid",
                     help="Your user id (same as used for login)")
    cli.add_argument("--cardid",
                     action='append',
                     help="Last 4 digits of your card number (*)")
    cli.add_argument("--output", "-o",
                     action='append',
                     help="Output path (QIF)")
    cli.add_argument("--qif-account",
                     action='append',
                     help="Default QIF account name (e.g. Aktiva:VISA) (*)")
    cli.add_argument("--from-date",
                     action='append',
                     help="Export transactions as of... (DD.MM.YYYY) (*)")
    cli.add_argument("--to-date",
                     action='append',
                     help="Export transactions until... (DD.MM.YYYY) (*)",
                     default=[date.today().strftime('%d.%m.%Y')])
    cli.add_argument("--raw", action="store_true",
                     help="Store the raw CSV file instead of QIF")
    cli.add_argument("--debug", action="store_true")
    cli.add_argument("--debug-dump", action="store_true")
    cli.add_argument("--debug-mechanize", action="store_true")
    cli.add_argument("--session-persistence-file")

    args = cli.parse_args()
    if not args.userid:
        cli.error("Please specify a valid user id")
    if len(args.cardid) == 0:
        cli.error("Please specify at least one valid card id")

    level = logging.INFO
    if args.debug:
        level = logging.DEBUG
    logging.basicConfig(level=level, format='%(message)s')

    def is_valid_date(date):
        return date and bool(re.match('^\d{1,2}\.\d{1,2}\.\d{2,5}\Z', date))

    def is_valid_dates(dates):
        return False not in [is_valid_date(date) for date in dates]

    from_date = args.from_date
    if len(args.cardid) == 1:
        while not is_valid_date(from_date[0]):
            from_date[0] = input("Start time: ")
    else:
        if len(args.from_date) != len(args.cardid) or not is_valid_dates(args.from_date):
            cli.error("Please specify exactly one valid start time for each card id")
    if len(args.to_date) not in [1, len(args.cardid)] or not is_valid_dates(args.to_date):
        cli.error("Please specify exactly one valid end time for each card id or none")
    if len(args.output) != len(args.cardid):
        cli.error("Please specify exactly one valid output path for each card id")
    if args.qif_account and len(args.qif_account) not in [0, len(args.cardid)]:
        cli.error("Please specify exactly one valid qif-account for each card id or none")

    fetcher = DkbScraper(record_html=args.debug_dump, session_persistence_file=args.session_persistence_file)

    if args.debug_mechanize:
        mechanize_logger = logging.getLogger("mechanize")
        mechanize_logger.addHandler(logging.StreamHandler(sys.stdout))
        mechanize_logger.setLevel(logging.INFO)
        #fetcher.br.set_debug_http(True)
        fetcher.br.set_debug_responses(True)
        #fetcher.br.set_debug_redirects(True)

    def get_pin_callback():
        if os.isatty(0):
            pin = ""
            while not pin.strip():
                pin = getpass('PIN: ')
            return pin
        else:
            return sys.stdin.read().strip()

    fetcher.login(args.userid, get_pin_callback)
    fetcher.credit_card_transactions_overview()
    for idx in range(len(args.cardid)):
        fetcher.select_transactions(args.cardid[idx], from_date[idx], args.to_date[idx] if idx < len(args.to_date) else args.to_date[0])
        csv_text = fetcher.get_transaction_csv()

        if args.raw:
            if args.output == '-':
                f = sys.stdout
            else:
                f = open(args.output[idx], 'wb')
            f.write(csv_text)
        else:
            cc_name = None
            if args.qif_account:
                cc_name = args.qif_account[idx]
            dkb2qif = DkbConverter(csv_text, cc_name=cc_name)
            dkb2qif.export_to(args.output[idx])
    fetcher.close()

# Testing
# =======
# python -m unittest dkb
# test_fetcher will fail unless you manually create test data, see below


class TestDkb(unittest.TestCase):
    def test_csv(self):
        text = open("tests/example.csv", "rb").read()
        c = DkbConverter(text)
        c.export_to("tests/example.qif")

    def test_fetcher(self):
        # Run with --debug-dump to create the necessary data for the tests.
        # This will record your actual dkb.de responses for local testing.
        f = DkbScraper(playback_html=True)
        f.BASEURL = "http://localhost:8000/loginform.html"
        f.br.set_debug_http(True)
        #f.br.set_debug_responses(True)
        f.br.set_debug_redirects(True)
        logger = logging.getLogger("mechanize")
        logger.addHandler(logging.StreamHandler(sys.stdout))
        logger.setLevel(logging.INFO)
        f.login("test", lambda: "1234")
        f.credit_card_transactions_overview()
        f.select_transactions("", "01.01.2013", "01.09.2013")
        print(f.get_transaction_csv())
        f.close()
