#!/usr/bin/env python2.7

import shlex, subprocess, time, sys

def makecommand(imapserver=None, adminuser=None, plevel="test", dryrun=True, user=None):
    imapsync_dir = "/opt/google-imap/"
    imapsync_cmd = imapsync_dir + "imapsync"
    cyrus_pf = imapsync_dir + "cyrus.pf"
    exclude_list = "'^Shared Folders|^mail/|^Junk$|^junk$|^JUNK$|^Spam$|^spam$|^SPAM$'"
    whitespace_cleanup = " --regextrans2 's/[ ]+/ /g' --regextrans2 's/\s+$//g' --regextrans2 's/\s+(?=\/)//g' --regextrans2 's/^\s+//g' --regextrans2 's/(?=\/)\s+//g'"
    folder_cases = " --regextrans2 's/^drafts$/[Gmail]\/Drafts/i' --regextrans2 's/^trash$/[Gmail]\/Trash/i' --regextrans2 's/^(sent|sent-mail)$/[Gmail]\/Sent Mail/i' --delete2foldersbutnot '^\[Gmail\]'"
    extra_opts = " --delete2 --delete2folders --fast"

    if dryrun:
        extra_opts = extra_opts + " --dry" 

    if plevel == "prod":
        google_pf = imapsync_dir + "google-prod.pf"
        google_domain = "pdx.edu"

    elif plevel == "test":
        google_pf = imapsync_dir + "google-test.pf"
        google_domain = "gtest.pdx.edu"

    else:
        raise Exception("Plevel must be test or prod.")

    command = imapsync_cmd + " --pidfile /tmp/imapsync-" + user + ".pid --host1 " + imapserver + " --port1 993 --user1 " + user + " --authuser1 " + adminuser + " --passfile1 " + cyrus_pf + " --host2 imap.gmail.com --port2 993 --user2 " + user + "@" + google_domain + " --passfile2 " + google_pf + " --ssl1 --ssl2 --maxsize 26214400 --authmech1 PLAIN --authmech2 XOAUTH -sep1 '/' --exclude " + exclude_list + folder_cases + whitespace_cleanup + extra_opts

    return command


if __name__ == "__main__":
    imapserver = "cyrus.psumail.pdx.edu"
    plevel = "test"
    adminuser = "cyradm"
    user = sys.argv[1]
    command = makecommand(imapserver, adminuser, plevel, True, user)
    print(command)
    syncprocess = subprocess.Popen(
        args=shlex.split(command)
        ,bufsize=-1
        ,close_fds=True
        ,stdout=None
        ,stderr=None
    )

    while (syncprocess.poll() == None):
       time.sleep(1) 

    if syncprocess.returncode == 0:
        print("ok")
    else:
        print("error_%d" % syncprocess.returncode)
