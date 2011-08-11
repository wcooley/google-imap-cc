import email
import imaplib
import xoauth
import base64
from pyparsing import Word, alphas, nums, printables, ZeroOrMore, ParseException

class imapstat:
    def __init__(self, imapserver=None, imapadmin=None, imappassword=None, gmaildomain=None, gmailsecret=None):
        """Sets parameters for the object: <imapserver>, <imapadmin>, <imappassword>, <gmaildomain> and <gmailsecret>."""
        self.imapserver = imapserver
        self.imapadmin = imapadmin
        self.imappassword = imappassword
        self.gmailserver = "imap.gmail.com"
        self.gmaildomain = gmaildomain
        self.gmailsecret = gmailsecret


    def cyr_connect(self, user = None):
        """Establishes a cyrus connection for <user>."""
        authstring = "%s\x00%s\x00%s" % (user, self.imapadmin, self.imappassword)

        self.imap = imaplib.IMAP4_SSL(self.imapserver)
        self.imap.authenticate("PLAIN", lambda x: authstring)


    def gmail_connect(self, user = None):
        """Establishes a gmail connection for <user>."""
        xoconsumer = xoauth.OAuthEntity(self.gmaildomain, self.gmailsecret)
        xotoken = xoauth.OAuthEntity("","")
        xouser = "%s@%s" % (user, self.gmaildomain)

        authstring = xoauth.GenerateXOauthString(xoconsumer, xotoken, xouser, "imap", xouser, None, None)

        self.imap = imaplib.IMAP4_SSL(self.gmailserver)
        self.imap.authenticate("XOAUTH", lambda x: authstring)


    def disconnect(self):
        """Closes an established IMAP cyr_connection."""
        self.imap.logout()


    def mboxstat(self, mbox):
        """Sends an IMAP select against the named <mbox>, returning True if the command succeeds, False otherwise."""
        sele_ret, msgs_cnt = self.imap.select(mbox, readonly = True)

        if sele_ret == "OK":
            return True
        else: 
            return False


    def mboxdel(self, mbox):
        """Deletes a IMAP <mbox>, returning True if the command succeeds, False otherwise."""
        sele_ret, msgs_cnt = self.imap.delete(mbox)

        if sele_ret == "OK":
            return True
        else: 
            return False


    def parsequota(self, rawdata):
        """Takes the raw output from a IMAP getquotaroot command, like so:
        [['INBOX INBOX'], ['INBOX (STORAGE 151788 1000000)']]

        Returns a 2-tuple, (quota_used, quota)

        Example:

        >>> ims = imapstat()
        >>> good = [['INBOX INBOX'], ['INBOX (STORAGE 151788 1000000)']]
        >>> nullquota = [['INBOX INBOX'], ['INBOX ()']]
        >>> bad = [['INBOX INBOX'], ['YUCK INBOX (STORAGE 151788 1000000)']]
        >>> ims.parsequota(rawdata = good)
        (151788, 1000000)
        >>> ims.parsequota(rawdata = nullquota)
        (0, 0)
        >>> ims.parsequota(rawdata = bad)
        Traceback (most recent call last):
            ...
        Exception: Error parsing: ['YUCK INBOX (STORAGE 151788 1000000)']
        """
        rootname = Word(alphas)
        resource = Word(alphas)
        quota_used = Word(nums)
        quota = Word(nums)

        quota_form = rootname + '(' + ZeroOrMore(resource + quota_used + quota) + ')'

        quota_parse = quota_form.parseString

        try:
            parsed = quota_parse(rawdata[1][0])
            used_quota = int(parsed[3])
            quota = int(parsed[4])

        except IndexError:
            quota, used_quota = (0, 0)

        except:
            raise Exception("Error parsing: %s" % rawdata[1])

        return (used_quota, quota)


    def parsemboxlist(self, rawdata):
        """Takes the raw output from a IMAP list mailboxes command, like so:
        ['(\\Noinferiors) "/" "INBOX"', '(\\HasNoChildren) "/" "Drafts"']

        Notice how it's a list of quoted strings.

        For strings that contain string delimiting characters, the data is returned as a tuple.

        Returns a list of quoted mailbox names "INBOX", "Drafts"...

        Example:

        >>> ims = imapstat()
        >>> goodstrings = ['(\\Noinferiors) "/" "INBOX"', '(\\HasNoChildren) "/" "Drafts"', '(\\HasNoChildren) "/" "I Love Spam"', '(\\HasNoChildren) "/" "Notes"', '(\\HasNoChildren) "/" "Sent"', '(\\HasChildren) "/" "Trash"', '(\\HasNoChildren) "/" "Trash/Sent"', '(\\HasNoChildren) "/" "Trash/Sent Messages"', '(\\HasChildren) "/" "_CECS"', '(\\HasNoChildren) "/" "_CECS/Announce"', '(\\HasNoChildren) "/" "_CECS/Asynchronous"', '(\\HasNoChildren) "/" "_CECS/CAT"', '(\\HasNoChildren) "/" "_CECS/CS162"', '(\\HasNoChildren) "/" "_CECS/CS163"', '(\\HasNoChildren) "/" "_CECS/CS201"', '(\\HasNoChildren) "/" "_CECS/CS202"', '']
        >>> goodtuple = [('(\\HasNoChildren) "/" {34}', 'Other Users/hyndlatest/foo "quote"')]
        >>> bad = ['(\\Noinferiors) "INBOX"']
        >>> ims.parsemboxlist(goodstrings)
        ['', 'Trash/Sent Messages', '_CECS/Announce', '_CECS/Asynchronous', 'Drafts', '_CECS/CAT', 'Notes', '_CECS', '_CECS/CS163', '_CECS/CS162', '_CECS/CS202', '_CECS/CS201', 'INBOX', 'Trash/Sent', 'I Love Spam', 'Trash', 'Sent']
        >>> ims.parsemboxlist(goodtuple)
        ['Other Users/hyndlatest/foo "quote"']
        >>> ims.parsemboxlist(bad)
        Traceback (most recent call last):
            ...
        Exception: Error parsing string (\Noinferiors) "INBOX"
        """
        flags = Word(alphas + '\\')
        root = Word(alphas + '/')
        mboxname = Word(printables + ' ')

        mbox_format = '(' + ZeroOrMore(flags) + ')' + '"' + root + '"' + mboxname

        mbox_parse = mbox_format.parseString

        parsed = set()

        for rawdatum in rawdata:
            if isinstance(rawdatum, str):
                try:
                    parsed.add(mbox_parse(rawdatum)[-1].strip('"'))

                except ParseException:
                    if rawdatum == "":
                        parsed.add("")

                    else:
                        raise Exception("Error parsing string %s" % str(rawdatum))

                except:
                    raise

            elif isinstance(rawdatum, tuple):
                try:
                    parsed.add(rawdatum[-1])

                except IndexError:
                    raise Exception("Error unpacking tuple %s" % str(rawdatum))

                except:
                    raise

            else:
                raise Exception("Error processing %s" % str(rawdatum))

        return list(parsed)


    def validatemboxnames(self, mboxes):
        """Takes a list of mailboxes as <mboxes> and searches for Google no-nos:

        Leading whitespace (also before /),
        Trailing whitespace (also after /),
        More than one space,
        Mailboxes that are identical if treated case insensitively.

        Returns None if all is cool, or a dictionary with the symptom as the key and the mailboxes in a list.

        Example:
        >>> ims = imapstat()
        >>> mboxlist = ['Shared Folders/archive/abuse/SpamReports/SR_Pubnet', 'Shared Folders/dept/cie/postmaster', 'Shared Folders/dept/arc/survey', 'sent', 'Spacey  ', '_PSU/Announce', 'Shared Folders/dept/cie/vendors', '_PSU/Administrative', 'Shared Folders/archive/abuse/spamfeedback', 'Shared Folders/dept/cie/os-updates', 'Trash/Sent', 'Shared Folders/dept/cie/webmaster', 'Sent', '_CECS/Advising', 'SENT', 'Shared Folders/dept/cie/printmaster', 'Shared Folders/dept/cie/unixteam', 'Sent-mail', 'Shared Folders/dept/oit/collab-team', 'Shared Folders/dept/cie/monitoring', '_Vendors/EMC', '_CECS/CS163', '_CECS/CS201', '_CECS/Machine Shop', 'Shared Folders/dept/cie/svn/adminutils', 'Shared Folders/archive/abuse/SpamReports/SR_CECS', 'Other Users/wcooley/unix student worker apps', 'Whitespace ', '_OIT', 'Shared Folders/dept/cie/svn/puppet', 'Shared Folders/archive/abuse/SpamReports/SR_Pre2007', 'Shared Folders/archive/abuse/spamfeedback/Scomp', 'Shared Folders/dept/cie/ids', '_Vendors/Sophos', '_OIT/Alerts', '_CECS/CS162', 'sent-mail', '_RT', '_PSU/Viking Motorsports', 'Shared Folders/archive/abuse/SpamReports/SR_ResNet', 'Shared Folders/dept/cie/license', 'INBOX', '_OIT/Email Access', '_Vendors/Spamhaus', '_PSU', 'Shared Folders/dept/cie/accountmaint', 'SENT-MAIL', 'Shared Folders/archive/abuse/SpamReports/solutions', '_CECS/CAT', 'Shared Folders/dept/cie/backups', 'Trash', 'Shared Folders/dept/cie/openpkg', 'Shared Folders/dept/cie/logs', '_CECS/CS251', 'Shared Folders/dept/cie/cfengine', '_OIT/Student Worker Apps', 'Shared Folders/dept/cie/maillists', '_OIT/Google', 'Shared Folders/archive/abuse/SpamReports/SR_Campus', 'Shared Folders/dept/cie/svn/cfengine', '_Vendors/Iron Mountain', 'Shared Folders/dept/cie/cron', 'I Love Spam', 'Shared Folders/dept/cie/install', 'Shared Folders/dept/cie/root', '_Vendors/Dell', '_CECS', '_Vendors/Red Hat', 'Notes', 'Shared Folders/dept/cie/massmail', 'Shared Folders/dept/cie/svn/nagios', '_CECS/Announce', 'Shared Folders/archive/abuse/spamtrap', '_OIT/Unixteam', 'Shared Folders/dept/cie/cfengine/disabled', 'Slash /Bang ', 'Drafts', 'Shared Folders/dept/cie', '_Vendors', '_Vendors/Oracle', 'Shared Folders/dept/cie/ids/Isaac', '_CECS/Asynchronous', 'Shared Folders/archive/abuse/SpamReports', 'Shared Folders/dept/cie/requests', 'Shared Folders/dept/cie/svn/massmail', 'Shared Folders/dept/cie/webmaster/sslmaster', '_OIT/Administrative', 'Shared Folders/archive/abuse', '_CECS/CS202', 'Trash/Sent Messages']
        >>> ims.validatemboxnames(mboxlist)
        {'trailing space': ['Spacey  ', 'Whitespace ', 'Slash /Bang '], 'case collision': ['sent', 'Sent', 'SENT', 'Sent-mail', 'sent-mail', 'SENT-MAIL']}
        """
        case_check = dict()

        problems = dict()
        problems["leading space"] = []
        problems["trailing space"] = []
        problems["multiple spaces"] = []
        problems["case collision"] = []

        # Initial pass--find whitespace violations, and populate case match dict.
        for mbox in mboxes:
            if len(mbox) > 0:
                if mbox[0] == " " or mbox.find("/ ") != -1:
                    problems["leading space"].append(mbox)

                if mbox[-1] == " " or mbox.find(" /") != -1:
                    problems["trailing space"].append(mbox)

                if mbox.find("  ") != -1:
                    problems["multiple spaces"].append(mbox)

                if case_check.has_key(mbox.lower()):
                    case_check[mbox.lower()].append(mbox)

                else:
                    case_check[mbox.lower()] = [mbox]

        # Find case collisions.
        for mboxgroup in case_check.values():
            if len(mboxgroup) > 1:
                for mbox in mboxgroup:
                    problems["case collision"].append(mbox)

        ok = True

        # Remove non-matching reasons.
        for (reason, mboxgroup) in problems.items():
            if mboxgroup == []:
                problems.pop(reason)

            else:
                ok = False

        # Return accordingly.
        if ok == False:
            return(problems)

        else:
            return None


    def parseheader(self, headers):
        """Takes the raw output from an IMAP fetch of message headers in <headers>, and Returns a list of dicts, each dict corresponding to a separate message header.

        Example:

        >>> ims = imapstat()
        >>> good = [('121 (RFC822.HEADER {1563}', 'Return-Path: <user@dom.tld>\\r\\nReceived: from murder (server.sub.dom.tld [192.168.122.1])\\r\\n\\t by server05.mail.dom.tld (Cyrus v2.3.7-Invoca-RPM-2.3.7-7.el5_4.3) with LMTPSA\\r\\n\\t (version=TLSv1/SSLv3 cipher=AES256-SHA bits=256/256 verify=YES);\\r\\n\\t Thu, 26 May 2011 12:20:56 -0700\\r\\nX-Sieve: CMU Sieve 2.3\\r\\nReceived: from server.sub.dom.tld ([unix socket])\\r\\n\\t by mail.dom.tld (Cyrus v2.2.13) with LMTPA;\\r\\n\\t Thu, 26 May 2011 12:20:56 -0700\\r\\nReceived: from server.sub.dom.tld (server.sub.dom.tld [192.168.111.42])\\r\\n\\tby server.sub.dom.tld (8.14.1+/8.13.1) with ESMTP id p4QJKup5003028\\r\\n\\tfor <user@server.dom.tld>; Thu, 26 May 2011 12:20:56 -0700\\r\\nReceived: from server-06.sub.dom.tld (server-06.sub.dom.tld [192.168.120.172])\\r\\n\\t(authenticated bits=0)\\r\\n\\tby server.sub.dom.tld (8.13.8/8.13.1) with ESMTP id p4QJKtW8021832\\r\\n\\t(version=TLSv1/SSLv3 cipher=DHE-RSA-AES256-SHA bits=256 verify=NOT)\\r\\n\\tfor <user@dom.tld>; Thu, 26 May 2011 12:20:55 -0700\\r\\nReceived: from server.sub.dom.tld (server.sub.dom.tld\\r\\n [192.168.132.34]) by server.dom.tld (Horde Framework) with HTTP; Thu, 26\\r\\n May 2011 12:20:55 -0700\\r\\nMessage-ID: <20110526122055.97746emge2ucfk7r@server.dom.tld>\\r\\nDate: Thu, 26 May 2011 12:20:55 -0700\\r\\nFrom: user@dom.tld\\r\\nTo: user@dom.tld\\r\\nSubject: Test Email\\r\\nMIME-Version: 1.0\\r\\nContent-Type: text/plain;\\r\\n charset=ISO-8859-1;\\r\\n DelSp="Yes";\\r\\n format="flowed"\\r\\nContent-Disposition: inline\\r\\nContent-Transfer-Encoding: 7bit\\r\\nUser-Agent: Dynamic Internet Messaging Program (DIMP) H3 (1.1.4)\\r\\nX-Scanned-By: MIMEDefang 2.71 on 192.168.111.42\\r\\n\\r\\n'), ')']
        >>> ims.parseheader(good)
        [{'Received': 'from murder (server.sub.dom.tld [192.168.122.1])\\r\\n\\t by server05.mail.dom.tld (Cyrus v2.3.7-Invoca-RPM-2.3.7-7.el5_4.3) with LMTPSA\\r\\n\\t (version=TLSv1/SSLv3 cipher=AES256-SHA bits=256/256 verify=YES);\\r\\n\\t Thu, 26 May 2011 12:20:56 -0700', 'X-Sieve': 'CMU Sieve 2.3', 'From': 'user@dom.tld', 'Return-Path': '<user@dom.tld>', 'MIME-Version': '1.0', 'Content-Transfer-Encoding': '7bit', 'X-Scanned-By': 'MIMEDefang 2.71 on 192.168.111.42', 'User-Agent': 'Dynamic Internet Messaging Program (DIMP) H3 (1.1.4)', 'To': 'user@dom.tld', 'Date': 'Thu, 26 May 2011 12:20:55 -0700', 'Message-ID': '<20110526122055.97746emge2ucfk7r@server.dom.tld>', 'Content-Type': 'text/plain;\\r\\n charset=ISO-8859-1;\\r\\n DelSp="Yes";\\r\\n format="flowed"', 'Content-Disposition': 'inline', 'Subject': 'Test Email'}]
        """
        return [
            dict(email.message_from_string(x[1]))
            for x in headers if x != ')'
        ]


    def quotastat(self):
        """Returns the current user's IMAP quota as the tuple: (quota_used, quota)."""
        quot_ret, quota_raw = self.imap.getquotaroot("INBOX")

        if quot_ret == "OK":
            return self.parsequota(quota_raw)
        else:
            raise Exception("Server returned invalid quota data")


    def mboxlist(self):
        """Returns a verified (can we IMAP select it?) list of a user's mailboxes."""
        mbox_ret, mbox_raw = self.imap.list()
        subm_ret, subm_raw = self.imap.lsub()

        list_raw = (mbox_raw + subm_raw)

        if (mbox_ret, subm_ret) == ("OK", "OK"):
            mbox_list = self.parsemboxlist(list_raw)

        else:
            raise Exception("Server returned invalid response to list command")

        return [x for x in mbox_list if self.mboxstat(x)]


    def bigmessages(self, user, mbox_list, lower_bound):
        "For a given <user>, and a <mbox_list> of that user's mailboxes, searches each mailbox for messages that exceed <lower_bound> bytes in size, and returns the headers of those messages in a dict of lists, where each key resolves to a list of dicts, each of which contains a separate message header."""
        msg_list = dict()

        self.cyr_connect(user)
        
        for mbox in mbox_list:
            try:
                sele_ret, msgs_cnt = self.imap.select(mbox, readonly = True)
                srch_ret, indx_raw = self.imap.search(None, "(LARGER %d)" % lower_bound)
    
                msg_headers = list()
                msg_indexes = indx_raw[0].split()
                        
                if len(msg_indexes) > 0:
                    msg_ret, msg_headers = self.imap.fetch(
                        ",".join(msg_indexes)
                        ,"RFC822.HEADER"
                    )
                
                    if msg_ret == "OK":
                        msg_list[mbox] = self.parseheader(msg_headers)

            except:
                print "Error processing mailbox %s" % mbox

        self.disconnect()

        return msg_list


    def stat(self, user):
        """For a given <user>, returns a dict containing a list of that user's accessible mailboxes, that user's quota, and how much of that quota is currently used."""
        mbox_list = list()
        quota = 0
        quota_used = 0

        self.cyr_connect(user)
        quota_used, quota = self.quotastat()
        mbox_list = self.mboxlist()
        mbox_problems = self.validatemboxnames(mbox_list)
        self.disconnect()

        return {"mbox_list":mbox_list,"quota":quota,"quota_used":quota_used,"mbox_problems":mbox_problems}


if __name__ == "__main__":
    import doctest
    doctest.testmod()
