#!/usr/bin/env python2
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
import csv
import sys
import logging
import mechanize
from mechanize._response import closeable_response, response_seek_wrapper
from StringIO import StringIO

DEBUG = False

logger = logging.getLogger(__name__)

class DKBBrowser(mechanize.Browser):
    """
    DKBBrowser is a mechanize.Browser which automatically fixes an HTML
    coding problem in the dkb.de non-js website.
    Sadly, this code must access non-public interfaces of mechanize.
    Let's hope, this code is only temporarily necessary...
    """
    def open(self, *args, **kwargs):
        response = mechanize.Browser.open(self, *args, **kwargs)
        if not response or not response.get_data():
            return response
        html = response.get_data().replace("<![endif]-->",
            "<!--[endif]-->")
        patched_resp = closeable_response(StringIO(html), response._headers,
            response._url, response.code, response.msg)
        patched_resp = response_seek_wrapper(patched_resp)
        self.set_response(patched_resp)
        return patched_resp

class DkbScraper(object):
    BASEURL = "https://banking.dkb.de/dkb/-"

    def __init__(self):
        self.br = DKBBrowser()

    def login(self, userid, pin):
        """
        Create a new session by submitting the login form

        @param str userid
        @param str pin
        """
        logger.info("Starting login as user %s...", userid)
        br = self.br

        # we are not a spider, so let's ignore robots.txt...
        br.set_handle_robots(False)

        # Although we have to handle a meta refresh, we disable it here
        # since mechanize seems to be buggy and will be stuck in a
        # long (infinite?) sleep() call
        br.set_handle_refresh(False)

        br.open(self.BASEURL + '?$javascript=disabled')

        # select login form:
        br.form = list(br.forms())[1]

        br.set_all_readonly(False)
        br.form["j_username"] = userid
        br.form["j_password"] = pin
        br.form["browserName"] = "Firefox"
        br.form["browserVersion"] = "40"
        br.form["screenWidth"] = "1000"
        br.form["screenHeight"] = "800"
        br.form["osName"] = "Windows"
        br.submit()
        br.open(self.BASEURL + "?$javascript=disabled")

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
                form.find_control(name="slAllAccounts")
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
            radio_ctrl = form.find_control("searchPeriod")
        except Exception:
            raise RuntimeError("Unable to find search period radio box")

        all_transactions_item = None
        for item in radio_ctrl.items:
            for label in item.get_labels():
                if re.search('Alle Ums.*tze.*', label.text, re.I):
                    all_transactions_item = item
                    break

        if not all_transactions_item:
            raise RuntimeError(
                "Unable to find 'Alle Umsätze vom' radio box")

        form[radio_ctrl.name] = [all_transactions_item.name]

        try:
            from_item = form.find_control(name="postingDate")
        except Exception:
            raise RuntimeError("Unable to find 'ab' date field")

        from_item.value = from_date

        try:
            to_item = form.find_control(name="toPostingDate")
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
            cc_list = form.find_control(name="slAllAccounts")
        except Exception:
            raise RuntimeError("Unable to find credit card selection form")

        for item in cc_list.get_items():
            # find right credit card...
            for label in item.get_labels():
                pattern = r'\b\S{12}(?<=%s)\b' % re.escape(cardid)
                if re.search(pattern, label.text, re.I):
                    form[cc_list.name] = [item.name]
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
        return self.br.response().read()


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
        self.csv_text = (csv_text
            .decode(self.INPUT_CHARSET)
            .encode(self.OUTPUT_CHARSET))
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
            .replace('.', '') # 1.000 -> 1000
            .replace(',', '.')) # 0,83 -> 0.83

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
        for x in xrange(self.SKIP_LINES):
            reader.next()
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
                f.write(line + "\n")

if __name__ == '__main__':
    from getpass import getpass
    from argparse import ArgumentParser
    from datetime import date

    level = logging.INFO
    if DEBUG:
        level = logging.DEBUG
    logging.basicConfig(level=level, format='%(message)s')

    cli = ArgumentParser()
    cli.add_argument("--userid",
        help="Your user id (same as used for login)")
    cli.add_argument("--cardid",
        help="Last 4 digits of your card number")
    cli.add_argument("--output", "-o",
        help="Output path (QIF)")
    cli.add_argument("--qif-account",
        help="Default QIF account name (e.g. Aktiva:VISA)")
    cli.add_argument("--from-date",
        help="Export transactions as of... (DD.MM.YYYY)")
    cli.add_argument("--to-date",
        help="Export transactions until... (DD.MM.YYYY)",
        default=date.today().strftime('%d.%m.%Y'))
    cli.add_argument("--raw", action="store_true",
        help="Store the raw CSV file instead of QIF")

    args = cli.parse_args()
    if not args.userid:
        cli.error("Please specify a valid user id")
    if not args.cardid:
        cli.error("Please specify a valid card id")

    def is_valid_date(date):
        return date and bool(re.match('^\d{1,2}\.\d{1,2}\.\d{2,5}\Z', date))

    from_date = args.from_date
    while not is_valid_date(from_date):
        from_date = raw_input("Start time: ")
    if not is_valid_date(args.to_date):
        cli.error("Please specify a valid end time")
    if not args.output:
        cli.error("Please specify a valid output path")

    pin = ""
    import os
    if os.isatty(0):
        while not pin.strip():
            pin = getpass('PIN: ')
    else:
        pin = sys.stdin.read().strip()

    fetcher = DkbScraper()

    if DEBUG:
        logger = logging.getLogger("mechanize")
        logger.addHandler(logging.StreamHandler(sys.stdout))
        logger.setLevel(logging.INFO)
        #fetcher.br.set_debug_http(True)
        fetcher.br.set_debug_responses(True)
        #fetcher.br.set_debug_redirects(True)

    fetcher.login(args.userid, pin)
    fetcher.credit_card_transactions_overview()
    fetcher.select_transactions(args.cardid, from_date, args.to_date)
    csv_text = fetcher.get_transaction_csv()

    if args.raw:
        if args.output == '-':
            f = sys.stdout
        else:
            f = open(args.output, 'w')
        f.write(csv_text)
    else:
        dkb2qif = DkbConverter(csv_text, cc_name=args.qif_account)
        dkb2qif.export_to(args.output)

# Testing
# =======
# python -m unittest dkb
# test_fetcher will fail unless you manually create test data, see below

import unittest
class TestDkb(unittest.TestCase):
    def test_csv(self):
        text = open("tests/example.csv", "rb").read()
        c = DkbConverter(text)
        c.export_to("/tmp/output.qif")

    def test_fetcher(self):
        # sorry, this test relies on private data, but I'll point out
        # how to create it:
        # - download the login form to loginform.html
        # - download the page after login, modify loginform.html to
        #   POST to this file instead of the original URL
        # - do the same for
        #     * the 'Kreditkartenumsätze' page,
        #     * the page after selecting a card and a time frame
        # - set up some lightweight web server (such as lighttpd) to
        #   serve these files and change BASEURL below accordingly
        # - this is also helpful if the web site changes and you have
        #   to adapt this code
        f = DkbScraper()
        f.BASEURL = "http://localhost:8000/loginform.html"
        f.br.set_debug_http(True)
        #f.br.set_debug_responses(True)
        f.br.set_debug_redirects(True)
        logger = logging.getLogger("mechanize")
        logger.addHandler(logging.StreamHandler(sys.stdout))
        logger.setLevel(logging.INFO)
        f.login("test", "1234")
        f.credit_card_transactions_overview()
        f.select_transactions("5678", "01.01.2013", "01.09.2013")
        print(f.get_transaction_csv())
