#!/usr/bin/env python
""" Archive Emails from a Gmail account """

############################################################################
#    Copyright (C) 2008 by Michael Goerz                                   #
#    http://www.physik.fu-berlin.de/~goerz                                 #
#                                                                          #
#    This program is free software; you can redistribute it and/or modify  #
#    it under the terms of the GNU General Public License as published by  #
#    the Free Software Foundation; either version 3 of the License, or     #
#    (at your option) any later version.                                   #
#                                                                          #
#    This program is distributed in the hope that it will be useful,       #
#    but WITHOUT ANY WARRANTY; without even the implied warranty of        #
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the         #
#    GNU General Public License for more details.                          #
#                                                                          #
#    You should have received a copy of the GNU General Public License     #
#    along with this program; if not, write to the                         #
#    Free Software Foundation, Inc.,                                       #
#    59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.             #
############################################################################

import sys
from optparse import OptionParser
import libgmail
from mailbox import mbox, mboxMessage
from cStringIO import StringIO
from time import sleep


def main(mboxfile, threadsfile=None, labelsfile=None, username=None, 
         password=None, verbose=False, label=None, delete=False, 
         msg_delay=0, thread_delay=0, skip_thread_delay = 0, 
         nodownload=False, query = None):
    """ Archive Emails from Gmail to an mbox """

    if username is None:
        username = raw_input("Gmail account name: ")
        
    if password is None:
        from getpass import getpass
        password = getpass("Password: ")

    if username.endswith("\n"): username = username[:-1]
    if password.endswith("\n"): password = password[:-1]

    ga = libgmail.GmailAccount(username, password)

    if verbose: print "\nPlease wait, logging in..."

    try:
        ga.login()
    except libgmail.GmailLoginFailure:
        print "\nLogin failed. (Wrong username/password?)"
    else:
        if verbose: print "Log in successful.\n"

        searches = libgmail.STANDARD_FOLDERS + ga.getLabelNames()

        while label is None and query is None:
            print "Select folder or label to archive: (Ctrl-C to exit)"

            for optionId, optionName in enumerate(searches):
                print "  %d. %s" % (optionId, optionName)
            print "  %d. %s" % (len(searches), "QUERY")

            try:
                label = searches[int(raw_input("Choice: "))]
            except ValueError:
                print "Please select a folder or label by typing in the " \
                      "number in front of it."
                raw_input()
                label = None
            except IndexError:
                query = raw_input("Query: ")
            except (KeyboardInterrupt, EOFError):
                print ""
                return 1
            print

        if verbose: 
            if label is not None:
                print "Selected folder: %s" % label
            else:
                print "Selected query: %s" % query

        if query is None:
            if label in libgmail.STANDARD_FOLDERS:
                result = ga.getMessagesByFolder(label, True)
            else:
                result = ga.getMessagesByLabel(label, True)
        else:
            result = ga.getMessagesByQuery(query, True)

        if len(result):
            archive_mbox = mbox(mboxfile)
            archive_mbox.lock()
            gmail_ids_in_mbox = {}
            # get dict of gmail_ids already in mbox
            for i, mbox_msg in enumerate(archive_mbox):
                gmail_ids_in_mbox[mbox_msg['X-GmailID']] = i
            labels_fh = StringIO()
            if labelsfile is not None:
                labels_fh = open(labelsfile, "w")
            threads_fh = StringIO()
            if threadsfile is not None:
                threads_fh = open(threadsfile, "w")
            gmail_ids = []
            try:
                for i, thread in enumerate(result):
                    if verbose: 
                        print "\nThread ID: ", thread.id, " LEN ", \
                              len(thread)
                    labels_fh.write("%s: %s\n" 
                                   % (thread.id, thread.getLabels()))
                    gmail_ids_in_thread = []
                    local_thread_delay = thread_delay
                    for gmail_msg in thread:
                        if verbose:
                            print "  ", gmail_msg.id, gmail_msg.number
                        gmail_ids_in_thread.append(str(gmail_msg.id))
                        gmail_ids.append(str(gmail_msg.id))
                        if nodownload:
                            if verbose: print "    skipped (no download)"
                            local_thread_delay = skip_thread_delay
                            continue
                        if gmail_ids_in_mbox.has_key(str(gmail_msg.id)):
                            if verbose: print "    skipped"
                            local_thread_delay = skip_thread_delay
                            continue # skip messages already in mbox
                        mbox_msg = mboxMessage(gmail_msg.source)
                        mbox_msg.add_header("X-GmailID", 
                                            gmail_msg.id.encode('ascii'))
                        archive_mbox.add(mbox_msg)
                        sleep(msg_delay)
                    threads_fh.write("%s\n" % gmail_ids_in_thread)
                    sleep(local_thread_delay)
                if delete:
                    for gmail_id in gmail_ids_in_mbox.keys():
                        if gmail_id not in gmail_ids:
                            mbox_id = gmail_ids_in_mbox[gmail_id]
                            if verbose: 
                                print "Delete id %s from mbox" % mbox_id
                            archive_mbox.remove(mbox_id)


            except KeyboardInterrupt:
                print "Keyboard Interrrupt"
            finally:
                if verbose: print "Flushing and closing mbox file"
                threads_fh.close()
                labels_fh.close()
                archive_mbox.close()
        else:
            if label is not None:
                print "No threads found in `%s`." % label
            else:
                print "No threads found in query `%s`." % query

    if verbose: print "\n\nDone."
    

if __name__ == "__main__":
    arg_parser = OptionParser(
    usage = "gmail_archive.py [options] MBOXFILE", description = __doc__)
        
    arg_parser.add_option('--username', action='store', type=str, 
                          dest='username', 
                          help="Gmail Username. If not specified, the program "
                          "will ask for it.")
    arg_parser.add_option('--password', action='store', type=str, 
                          dest='password', 
                          help="Gmail Password. If not specified, the program "
                          "will ask for it")
    arg_parser.add_option('--authfile', action='store', type=str, 
                          dest='authfile', 
                          help="File containing the username on the first "
                          "line, and the password in the second. This is "
                          "a replacement for specifying --username and "
                          "--password")
    arg_parser.add_option('--threadsfile', action='store', type=str, 
                          dest='threadsfile', 
                          help="File for storing thread information. If not "
                          "specified, no thread information will be stored.")
    arg_parser.add_option('--labelsfile', action='store', type=str, 
                          dest='labelsfile', 
                          help="File for storing label information. If not "
                          "specified, no label information will be stored.")
    arg_parser.add_option('--label', action='store', type=str, 
                          dest='label', 
                          help="Label or Folder to archive. If not specified, "
                          "the program will ask for it")
    arg_parser.add_option('--query', action='store', type=str, 
                          dest='query', 
                          help="Search to archive. This is an alternative to"
                          "specifying a --label. Only the messages that match"
                          "the search are archived. The --lable option takes "
                          "preference over --query")
    arg_parser.add_option('--msg_delay', action='store', type=int, 
                          dest='msg_delay', default=0,
                          help="Number of seconds to wait between accessing "
                          "messages. This and the following delays may "
                          "hopefully prevent you being locked out of your "
                          "account")
    arg_parser.add_option('--thread_delay', action='store', type=int, 
                          dest='thread_delay', default=0,
                          help="Number of seconds to wait between accessing "
                          "threads.")
    arg_parser.add_option('--skip_thread_delay', action='store', type=int, 
                          dest='skip_thread_delay', default=0,
                          help="Number of seconds to wait between accessing "
                          "threads that are not downloaded")
    arg_parser.add_option('--verbose', action='store_true', dest='verbose',
                          default=False, help="Print status messages")
    arg_parser.add_option('--delete', action='store_true', 
                          dest='delete',
                          default=False, help="Delete archived emails that "
                          "are no longer on the server")
    arg_parser.add_option('--nodownload', action='store_true', 
                          dest='nodownload',
                          default=False, help="Do not store any messages in "
                          "the mbox (behave like if the message was already "
                          "present there).")
    options, args = arg_parser.parse_args(sys.argv)

    try:
        mboxfile = args[1]
    except IndexError:
        print >> sys.stderr, "You did not provide enough parameters"
        arg_parser.print_help()
        sys.exit(1)
    if options.authfile is not None:
        auth_fh = open(options.authfile)
        options.username = auth_fh.readline()
        options.password = auth_fh.readline()
        auth_fh.close()

    main(mboxfile, options.threadsfile, options.labelsfile, options.username, 
         options.password, options.verbose, options.label, options.delete, 
         options.msg_delay, options.thread_delay, options.skip_thread_delay, 
         options.nodownload, options.query)
