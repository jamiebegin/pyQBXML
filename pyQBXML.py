import datetime, time
import sys
import os
import re
import string
import httplib

from lxml import etree
from decimal import Decimal

class QBInvoices(object):
    def __init__(self):
        self.invoices = []

    def __iter__(self):
        return iter(self.invoices)

    def __len__(self):
        return len(self.invoices)

    def add(self, invoice):
        self.invoices.append(invoice)

    def getByRequestID(self, request_id):
        invoice = [x for x in self.invoices if x.request_id == request_id] or None
        return invoice[0]

class QBInvoice(object):
    def __init__(self, invoice_date, customer_id, memo=None, terms=None, due_date=None, request_id=None, auto_create_items=False):
        self.customer_id = customer_id
        self.invoice_date = invoice_date
        self.memo = memo
        self.line_items = QBLineItems()

        if terms is None:
            self.terms = "Net 30"

        if due_date is None:
            today = datetime.date.today()
            net30 = datetime.timedelta(days=30)
            self.due_date = today + net30

        if request_id:
            self.request_id = request_id
        else:
            self.request_id = ''.join(random.choice(string.letters) for i in xrange(16))

        self.time_created = None
        self.time_modified = None
        self.customer_name = None
        self.is_paid = None
        self.auto_create_items = auto_create_items

    def addLineItem(self, qty, fullname, description, rate, item_type=None, account=None):
        if self.auto_create_items and not item_type:
            raise QBOEItemError("Item Type must be specified for each line item when using item autocreation.")
        if self.auto_create_items and not account:
            raise QBOEItemError("The Item Account must be specified for each line item when using item autocreation.")

        item = QBLineItem(qty, fullname, description, rate, item_type, account)
        self.line_items.add(item)

    def getLineItemByName(self, item_name):
        line_item = [x for x in self.line_items if x.fullname == item_name] or None
        return line_item[0]

    def serialize(self, request_id=None):
        el_invoice = etree.Element("InvoiceAdd")

        el_custref = etree.SubElement(el_invoice, "CustomerRef")

        el_date = etree.SubElement(el_invoice, "TxnDate")
        el_date.text = str(self.invoice_date)
        el_termsref = etree.SubElement(el_invoice, "TermsRef")
        el_terms = etree.SubElement(el_termsref, "FullName")
        el_terms.text = self.terms
        el_duedate = etree.SubElement(el_invoice, "DueDate")
        el_duedate.text = str(self.due_date)

        if self.memo:
            el_memo = etree.SubElement(el_invoice, "Memo")
            el_memo.text = str(self.memo)

        el_cust = etree.SubElement(el_custref,'ListID')
        el_cust.text = str(self.customer_id)

        for line_item in self.line_items:
            el_lineitem = etree.SubElement(el_invoice, "InvoiceLineAdd")

            el_itemref = etree.SubElement(el_lineitem, "ItemRef")
            el_fullname = etree.SubElement(el_itemref, "FullName")
            el_fullname.text = str(line_item.fullname)

            el_desc = etree.SubElement(el_lineitem, "Desc")
            el_desc.text = str(line_item.description)

            el_qty = etree.SubElement(el_lineitem, "Quantity")
            if line_item.qty == int(line_item.qty):
                el_qty.text = "%d" % (line_item.qty)
            else:
                el_qty.text = "%.2f" % (line_item.qty)
            el_rate = etree.SubElement(el_lineitem, "Rate")
            el_rate.text = "%.2f" % (line_item.rate)
        return el_invoice

class QBItemType(object):
    SERVICE = 1
    INVENTORY = 2
    NONINVENTORY = 3
    OTHER_CHARGE = 4
    GROUP = 5
    FIXED_ASSET = 6
    DISCOUNT = 7
    PAYMENT = 8
    SALES_TAX = 9
    SALES_TAX_GROUP = 10
    SUBTOTAL = 11

class QBLineItems(object):
    def __init__(self):
        self.line_items = []

    def __iter__(self):
        return iter(self.line_items)

    def __len__(self):
        return len(self.line_items)

    def add(self, line_item):
        self.line_items.append(line_item)

class QBLineItem(object):
    def __init__(self, qty, fullname, description, rate, item_type=None, account=None):
        self.fullname = fullname
        self.description = description
        self.rate = rate
        self.qty = qty

        if item_type:
            enum = [(i,v) for v,i in QBItemType.__dict__.iteritems() if v[:2] != "__"]
            try:
                if int(item_type) in [i for i,v in enum]:
                    self.item_type = item_type
                else:
                    raise ValueError
            except ValueError:
                types = "QBItemType."+", QBItemType.".join([v for i,v in enum]) + ")"
                raise QBOEItemError("Invalid Item Type specified. (Hint: Allowed types are: " + types)
        self.account = account

class QBCustomers(object):
    def __init__(self):
        self.customers = []

    def __iter__(self):
        return iter(self.customers)

    def __len__(self):
        return len(self.customers)

    def add(self, customer):
        self.customers.append(customer)

class QBCustomer(object):
    def __init__(self, list_id, name):
        self.list_id = list_id
        self.name = name
        self.time_created = None
        self.time_modified = None
        self.first_name = None
        self.last_name = None
        self.full_name = None
        self.edit_sequence = None
        self.sublevel = None
        self.print_as = None
        self.phone = None
        self.email = None
        self.delvery_method = None
        self.balance = Decimal('0.00')
        self.total_balance = Decimal('0.00')
        self.is_statement_with_parent = None
        self.bill_address = QBAddress()

class QBAddress(object):
    def __init__(self):
        self.address1 = None
        self.address2 = None
        self.city = None
        self.state = None
        self.postal_code = None

class QBOEError(BaseException):
    def __init__(self, err_msg):
        self.err_msg = err_msg

    def __str__(self):
        return "%s" % self.err_msg

class QBOEItemError(QBOEError):
    pass

class QBXMLError(QBOEError):
    def __init__(self, err_code, err_msg):
        self.err_code = err_code
        self.err_msg = err_msg

    def __str__(self):
        return "%s (qbXML statusCode: %d)" % (self.err_msg, self.err_code)

class QBOEHTTPError(QBOEError, httplib.HTTPException):
    def __init__(self, httplib_ex, err_msg):
        self.httplib_ex = httplib_ex
        self.err_msg = err_msg

    def __str__(self):
        print ": %s (%s %s)" % (self.err_msg, self.httplib_ex.__class__.__name__, self.httplib_ex)
        raise self.httplib_ex

class QBOE(object):
    def __init__(self, api_url, key_file, cert_file, app_name, app_id, app_ver, conn_ticket, https_timeout=60, debug=False):
        self.api_url = api_url
        self.key_file = key_file
        self.cert_file = cert_file
        self.app_name = app_name
        self.app_name_id = app_id
        self.app_name_ver = app_ver
        self.conn_ticket = conn_ticket
        self.https_timeout = https_timeout

        self.debug = debug
        self.__session_ticket = None

        self.invoices = QBInvoices()

    def __makeQBXMLReq(self, data):
        """
        Add the session authentication information to the specificed qbXML document
        in preperation for submission to Quickbooks.
        """
        root = etree.Element("QBXML")
        tree = etree.ElementTree(root)
        root.addprevious(etree.ProcessingInstruction ('qbxml', 'version="6.0"'))

        el_signon = etree.SubElement(root, "SignonMsgsRq")

        el_signon_ticket = etree.SubElement(el_signon, 'SignonTicketRq')

        el_datetime = etree.SubElement(el_signon_ticket,'ClientDateTime')
        el_datetime.text = self.__getXMLDatetime()

        el_session_ticket = etree.SubElement(el_signon_ticket,'SessionTicket')
        el_session_ticket.text = self.__session_ticket

        el_lang = etree.SubElement(el_signon_ticket,'Language')
        el_lang.text = 'English'

        el_app_id = etree.SubElement(el_signon_ticket,'AppID')
        el_app_id.text = self.app_name_id

        el_ver = etree.SubElement(el_signon_ticket,'AppVer')
        el_ver.text = self.app_name_ver

        root.append(data)

        if self.debug:
            print etree.tostring(tree, pretty_print=True, encoding="utf-8", xml_declaration=True)
        return tree

    def __submitQBXMLReq(self, xmldoc, recursing=False):
        """
        Send the specified XML document to the Quickbooks QPI for processing via a HTTPS POST.
        """

        if not self.__session_ticket and not recursing:
            # We're not logged in (no ticket) and this is the first time we're through here.
            req = self.__makeSignInReq()
            signin_response = self.__submitQBXMLReq(xmldoc=req, recursing=True)

            assert self.__parseLoginResponse(signin_response) == True, "Unable to parese login response."

            # Logged in ok, so we now update the session ticket in the original request
            ticket_el = xmldoc.xpath("/QBXML/SignonMsgsRq/SignonTicketRq/SessionTicket")[0]
            ticket_el.text = self.__session_ticket

        headers = {"Content-type": "application/x-qbxml"}
        host = self.api_url.split("/")[0]
        path = "/" + "/".join(self.api_url.split("/")[1:])

        h = httplib.HTTPSConnection(
                            host=host
                            ,port=443
                            ,key_file=self.key_file
                            ,cert_file=self.cert_file
                            ,timeout=self.https_timeout)
        if self.debug:
            h.debuglevel=1

        data = etree.tostring(xmldoc
                                ,pretty_print=False
                                ,encoding="utf-8"
                                ,xml_declaration=True)

        try:
            h.request('POST', path, data, headers)
            resp = h.getresponse()
            data = resp.read()
            if not resp.status == 200:
                raise QBOEHTTPError(None, "Invalid response received from QBOE. Response: %d %s" % (resp.status, resp.reason))
        except httplib.ssl.SSLError as ex:
            self.__checkCerts()
            try:
                httplib_err_num = int(re.search(r'error:([0-9A-F]+):', ex.args[1], re.I|re.M).group(1).encode("hex"))
            except (AttributeError, ValueError):
                httplib_err_num = None

            if httplib_err_num == 3134304230303039:
                raise QBOEHTTPError(ex, "There appears to be a problem with the specified private key file: '%s'."
                                        " (Hint: Is the first line of this file '-----BEGIN RSA PRIVATE KEY-----'?)\n\n" \
                                        % self.key_file)
            elif httplib_err_num == 3134304443303039:
                raise QBOEHTTPError(ex, "There appears to be a problem with the specified certificate file: '%s'."
                                        " (Hint: Is the first line of this file '-----BEGIN CERTIFICATE-----'?)\n\n" \
                                        % self.cert_file)
            elif httplib_err_num == 3134303934343142:
                raise QBOEHTTPError(ex, "The specified certificate ('%s') and key ('%s') don't match or are corrupted.\n\n" \
                                        % (self.cert_file, self.key_file))
            else:
                raise
        finally:
            h.close()

        return etree.XML(data)

    def __makeSignInReq(self):
        """
        Generate the XML document that contains the sign-in request for the Quickbooks API
        """
        root = etree.Element("QBXML")
        tree = etree.ElementTree(root)
        root.addprevious(etree.ProcessingInstruction ('qbxml', 'version="6.0"'))

        el_signon = etree.SubElement(root, "SignonMsgsRq")
        el_app_cert = etree.SubElement(el_signon, 'SignonAppCertRq')

        el_datetime = etree.SubElement(el_app_cert,'ClientDateTime')
        el_datetime.text = self.__getXMLDatetime()

        el_app = etree.SubElement(el_app_cert,'ApplicationLogin')
        el_app.text = self.app_name

        el_ticket = etree.SubElement(el_app_cert,'ConnectionTicket')
        el_ticket.text = self.conn_ticket

        el_lang = etree.SubElement(el_app_cert,'Language')
        el_lang.text = 'English'

        el_app_id = etree.SubElement(el_app_cert,'AppID')
        el_app_id.text = self.app_name_id

        el_ver = etree.SubElement(el_app_cert,'AppVer')
        el_ver.text = self.app_name_ver

        if self.debug:
            print etree.tostring(tree, pretty_print=True, encoding="utf-8", xml_declaration=True)
        return tree

    def __parseLoginResponse(self, root):
        """
        Evaluate the response to the sign-in attempt.  Check for errors and parse out the session ticket
        that is included with further requests.
        """
        ticket = None
        msgs = root.xpath("//SignonMsgsRs/*")
        for msg in msgs:
            if msg.tag == "SignonAppCertRs" or msg.tag == "SignonTicketRs":
                if msg.attrib['statusSeverity'] == 'ERROR':
                    raise QBXMLError(int(msg.attrib['statusCode']), msg.attrib['statusMessage'])
                elif msg.attrib['statusSeverity'] == 'INFO':
                    if msg.xpath('SessionTicket'):
                        ticket = msg.xpath('SessionTicket')[0].text
        if ticket:
            self.__session_ticket = ticket
            return True
        else:
            raise QBXMLError(-1, "Expected to a receive session ticket or error but got neither. Cannot login.")
            return False

    def __checkCerts(self):
        """
        If a SSL exception occurs, check to ensure that the user-specifiec certificate and key files are readable.
        """
        if not os.access(self.key_file, os.R_OK):
            if not os.path.exists(self.key_file):
                raise QBOEError('The specified SSL key file ("%s") does not exist.' % self.key_file)
            else:
                raise QBOEError('The specified SSL key file ("%s") exists but is is not readable.' % self.key_file)
        elif not os.access(self.cert_file, os.R_OK):
            if not os.path.exists(self.cert_file):
                raise QBOEError('The specified SSL certificate file ("%s") does not exist.' % self.cert_file)
            else:
                raise QBOEError('The specified SSL certificate file ("%s") exists but is is not readable.' % self.cert_file)

    def __getXMLDatetime(self, stamp=datetime.datetime.now()):
        """
        Convert a Python DateTime object into the string format used by qbXML.
        """
        s = stamp.strftime("%Y-%m-%dT%H:%M:%S")
        return s

    def __XMLToDatetime(self, stamp):
        """
        Convert qbXML DateTime string into Python DateTime object.
        """
        s = datetime.datetime(*time.strptime(stamp,"%Y-%m-%dT%H:%M:%S")[0:5])
        return s

    def __XMLToDate(self, stamp):
        """
        Convert qbXML DateTime string into Python Date object.
        """
        d = [int(x) for x in stamp.split("-")]
        s = datetime.date(*d)
        return s

    def addInvoice(self, invoice):
        """
        Add an individual invoice to the submission queue. putInvoices() is later called to submit
        this queue to the Quickbooks API.
        """
        self.invoices.add(invoice)

    def putInvoices(self, specific_invoices=None):
        """
        Serialize the invoice queue into a qbXML document and submit the batch to Quickbooks for posting.
        """
        root = etree.Element('QBXMLMsgsRq')
        root.set("onError", "continueOnError")

        if not specific_invoices:
            specific_invoices = self.invoices
        for invoice in specific_invoices:
            el = etree.SubElement(root,'InvoiceAddRq')
            el.set("requestID", str(invoice.request_id) or "")
            el.append(invoice.serialize())

        res = self.__makeQBXMLReq(root)
        xmldoc = self.__submitQBXMLReq(res)

        if self.debug:
            print etree.tostring(xmldoc, pretty_print=True, encoding="utf-8", xml_declaration=True)

        line_items_to_create = QBLineItems()
        invoices_to_redo = []
        msgs = xmldoc.xpath('/QBXML/QBXMLMsgsRs/InvoiceAddRs')
        for msg in msgs:
            if msg.get('statusSeverity') == 'Error':
                    status_code = int(msg.attrib['statusCode'])
                    if status_code == 3140:
                        # Trying to add an item that doesn't exist in QB
                        this_invoice = self.invoices.getByRequestID(msg.attrib['requestID'])
                        if this_invoice.auto_create_items:
                            try:
                                item_name = re.match(r'Invalid reference to ItemList: (.+) in ItemRef', msg.attrib['statusMessage']).group(1)
                            except AttributeError:
                                raise QBOEItemError("Cannot add a line item with an empty item name to an invoice.")
                            li = this_invoice.getLineItemByName(item_name)
                            line_items_to_create.add(li)
                            invoices_to_redo.append(this_invoice)
                        else:
                            raise QBOEItemError("Invoice contains at least one line item that does not exist in QBOE. [Request ID: '%s']"
                                        " (Hint: Set the 'auto_create_items=True' QBInvoice property to automatically create items missing from QBOE." \
                                        " Refer to the online documentation for more info.)" % msg.attrib['requestID'])

                    else:
                        raise QBXMLError(status_code, msg.attrib['statusMessage'])

        invoices = {}
        if len(line_items_to_create) > 0:
            self.__createItems(line_items_to_create)

            redone = self.putInvoices(specific_invoices=invoices_to_redo)
            invoices.update(redone)

        for invoice in xmldoc.xpath('/QBXML/QBXMLMsgsRs/InvoiceAddRs'):
            request_id = invoice.get("requestID")
            if len(request_id) > 0 and request_id not in invoices:
                invoice_num = invoice.xpath('InvoiceRet/RefNumber')[0].text
                invoices[request_id] = invoice_num

        return invoices

    def __createItems(self, line_items):
        for li in line_items:
            if li.item_type == QBItemType.SERVICE:
                self.addServiceItem(li.fullname, li.description, li.rate, li.account)
            else:
                raise QBOEItemError("Currently only QBItemType.SERVICE items may be added or modified. Please note that"
                                    " this is a limitation of QBOE's subset of qbXML and not pyQBOE itself.")

    def addServiceItem(self, item_name, description, rate, account):
        root = etree.Element('QBXMLMsgsRq')
        root.set("onError", "continueOnError")

        el_base = etree.SubElement(root,'ItemServiceAddRq')
        el_base.set("requestID", "")

        el_isa = etree.SubElement(el_base, "ItemServiceAdd")

        el_name = etree.SubElement(el_isa, "Name")
        el_name.text = str(item_name)

        el_sop = etree.SubElement(el_isa, "SalesOrPurchase")

        el_desc = etree.SubElement(el_sop, "Desc")
        el_desc.text = str(description)

        el_price = etree.SubElement(el_sop, "Price")
        el_price.text = "%.2f" % rate

        el_ar = etree.SubElement(el_sop, "AccountRef")
        el_an = etree.SubElement(el_ar, "FullName")
        el_an.text = str(account)

        res = self.__makeQBXMLReq(root)
        xmldoc = self.__submitQBXMLReq(res)

        if self.debug:
            print etree.tostring(xmldoc, pretty_print=True, encoding="utf-8", xml_declaration=True)

    def getInvoices(self):
        """
        Retrieve the list of invoices from Quickbooks.
        """
        root = etree.Element('QBXMLMsgsRq')
        root.set("onError", "continueOnError")
        el = etree.SubElement(root,'InvoiceQueryRq')

        res = self.__makeQBXMLReq(root)
        xml = self.__submitQBXMLReq(res)

        invoices = QBInvoices()
        for invoice in xml.xpath("QBXMLMsgsRs/InvoiceQueryRs/*"):
            if len(invoice.xpath('TxnDate')) > 0 and len(invoice.xpath('CustomerRef/ListID')) > 0:
                i = QBInvoice( invoice_date =  self.__XMLToDate(invoice.xpath('TxnDate')[0].text)
                                ,customer_id = invoice.xpath('CustomerRef/ListID')[0].text)

                if len(invoice.xpath('TimeCreated')) > 0:
                    t = self.__XMLToDatetime(invoice.xpath('TimeCreated')[0].text)
                    i.time_created = t

                if len(invoice.xpath('TimeModified')) > 0:
                    t = self.__XMLToDatetime(invoice.xpath('TimeModified')[0].text)
                    i.time_modified = t

                if len(invoice.xpath('IsPaid')) > 0:
                    x = invoice.xpath('IsPaid')[0].text
                    if x.lower() == "true":
                        i.is_paid = True
                    else:
                        i.is_paid = False

                if len(invoice.xpath('CustomerRef/FullName')) > 0:
                    i.customer_name = invoice.xpath('CustomerRef/FullName')[0].text

                for line_item in invoice.xpath("InvoiceLineRet"):
                    if len(line_item.xpath('ItemRef/FullName')) > 0:
                        fullname = line_item.xpath('ItemRef/FullName')[0].text
                    if len(line_item.xpath('Desc')) > 0:
                        description = line_item.xpath('Desc')[0].text
                    if len(line_item.xpath('Rate')) > 0:
                        rate = Decimal(line_item.xpath('Rate')[0].text)
                    if len(line_item.xpath('Quantity')) > 0:
                        qty = Decimal(line_item.xpath('Quantity')[0].text)
                    i.addLineItem(fullname, description, rate, qty)


                invoices.add(i)
        return invoices


    def getCustomers(self, request_id=''):
        """
        Retreive the list of customers from Quickbooks.
        """
        customers = QBCustomers()
        root = etree.Element('QBXMLMsgsRq')
        root.set("onError", "continueOnError")
        el = etree.SubElement(root,'CustomerQueryRq')
        el.set("requestID", str(request_id))

        res = self.__makeQBXMLReq(root)
        xml = self.__submitQBXMLReq(res)

        for customer in xml.xpath("QBXMLMsgsRs/CustomerQueryRs/*"):
            if len(customer.xpath('Name')) > 0 and len(customer.xpath('ListID')) > 0:
                c = QBCustomer( list_id = customer.xpath('ListID')[0].text
                                ,name = customer.xpath('Name')[0].text)

                if len(customer.xpath('TimeCreated')) > 0:
                    t = self.__XMLToDatetime(customer.xpath('TimeCreated')[0].text)
                    c.time_created = t

                if len(customer.xpath('TimeModified')) > 0:
                    t =  self.__XMLToDatetime(customer.xpath('TimeModified')[0].text)
                    c.time_modified = t

                if len(customer.xpath('FullName')) > 0:
                    c.full_name = customer.xpath('FullName')[0].text
                if len(customer.xpath('CompanyName')) > 0:
                    c.company_name = customer.xpath('CompanyName')[0].text
                if len(customer.xpath('PrintAs')) > 0:
                    c.print_as = customer.xpath('PrintAs')[0].text
                if len(customer.xpath('EditSequence')) > 0:
                    c.edit_sequence = customer.xpath('EditSequence')[0].text
                if len(customer.xpath('Sublevel')) > 0:
                    c.sublevel = customer.xpath('Sublevel')[0].text
                if len(customer.xpath('Phone')) > 0:
                    c.phone = customer.xpath('Phone')[0].text
                if len(customer.xpath('Email')) > 0:
                    c.email = customer.xpath('Email')[0].text

                if len(customer.xpath('Balance')) > 0:
                    c.balance = Decimal(customer.xpath('Balance')[0].text or '0.00')

                if len(customer.xpath('TotalBalance')) > 0:
                    c.total_balance = Decimal(customer.xpath('TotalBalance')[0].text or '0.00')

                if len(customer.xpath('DeliveryMethod')) > 0:
                    c.delivery_method = customer.xpath('DeliveryMethod')[0].text

                if len(customer.xpath('IsStatementWithParent')) > 0:
                    c.is_statement_with_parent = customer.xpath('IsStatementWithParent')[0].text

                if len(customer.xpath('BillAddress')) > 0:
                    for a in customer.xpath('BillAddress'):
                        addr = QBAddress()
                        if len(a.xpath('Addr1')) > 0:
                            addr.address1 = a.xpath('Addr1')[0].text
                        if len(a.xpath('Addr2')) > 0:
                            addr.address2 = a.xpath('Addr2')[0].text
                        if len(a.xpath('City')) > 0:
                            addr.city = a.xpath('City')[0].text
                        if len(a.xpath('State')) > 0:
                            addr.state = a.xpath('State')[0].text
                        if len(a.xpath('PostalCode')) > 0:
                            addr.postal_code = a.xpath('PostalCode')[0].text
                        c.bill_address = addr

                customers.add(c)
        return customers

if __name__ == '__main__':
    qb = QBOE( api_url = "webapps.quickbooks.com/j/AppGateway"
                ,key_file = "./my_key.pem"
                ,cert_file = "./my_cert.crt"
                ,conn_ticket = 'TGT-104-zH084yIDGkH4_r2DYUUcevQ'
                ,app_name = 'myapp.mydomain.com'
                ,app_id = '112734952'
                ,app_ver = '1'
                ,debug = False)



    customers = qb.getCustomers()
    for customer in customers:
        print customer.full_name, customer.phone, customer.bill_address.city


    request_ids = ["invoice1", "invoice2", "invoice3"]
    for request_id in request_ids:
        invoice = QBInvoice(     customer_id = 1
                                ,invoice_date = datetime.date.today()
                                ,memo = "** Autogenerated by pyQBOE on %s" % datetime.date.today()
                                ,request_id = request_id)

        invoice.addLineItem('', 'Rocket-powered sled, perfect for chasing roadrunners!', 800.00, 1)
        qb.addInvoice(invoice)

    results = qb.putInvoices()
    print results

