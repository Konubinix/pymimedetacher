#!/usr/bin/env python
# -*- coding: utf-8 -*-
import mailbox
import tempfile
import os
import optparse
import re
import chardet
import quopri
import base64
import hashlib

# Input path with (Courier) maildir data
PATH = os.path.expanduser('~/.mail')
# Output path to store the attachments
OUTPATH = os.path.expanduser('~/detachments/')+PATH.split(os.sep)[-1] # the idea is to have an output folder per account

parser = optparse.OptionParser()

parser.add_option('-i', '--input', action="store", dest="PATH",
                  help="input maildir path to parse", default=PATH)
parser.add_option('-o', '--output', action="store", dest="OUTPATH",
                  help="output path to store the attachments", default=OUTPATH)
parser.add_option('-d', '--delete-attachment', action="store_true",
                  dest="del_attach", help="delete the attachments", default=False)
parser.add_option('-s', '--save_attachment', action="store_true",
                  dest="save_attach", help="save the attachments", default=False)
parser.add_option('-v', '--verbose', action="store_true",
                  dest="verbose", help="verbose output", default=False)

options, args = parser.parse_args()

PATH   = os.path.abspath(os.path.expanduser(options.PATH))
OUTPATH = os.path.abspath(os.path.expanduser(options.OUTPATH))


def ensure_unicode(value):
    """Convert a string in unicode"""
    if isinstance(value, str):
        decoded = None
        charsets = ("latin-1", "utf-8", "ascii")
        for i, charset in enumerate(charsets):
            try:
                return value.decode(charset)
            except UnicodeDecodeError:
                if i == len(charsets) - 1:
                    raise
    else:
        return value


def decode(value):
    match = re.match(
        "^=\?(?P<encoding>[^?]+)\?(?P<method>q|b|B|Q)\?(?P<content>[^?]+)\?=(?P<rest>.*)$",
        value,
    )
    method_handlers = {
        "q": quopri.decodestring,
        "b": base64.decodestring,
    }
    if match:
        value = method_handlers[match.group("method").lower()](
            match.group("content").encode("utf-8")
        ).decode(match.group("encoding")) + match.group("rest")
    return value


print 'Options :'
print '%20s : %s' % ('Mailbox Path', PATH)
print '%20s : %s' % ('Output Path ', OUTPATH)
print '%20s : %s' % ('delete attachment', options.del_attach)
print '%20s : %s' % ('save attachment', options.save_attach)
print '%20s : %s' % ('verbose', options.save_attach)
# Useful links:
# - MIME structure :Parsing email using Python part 2,  http://blog.magiksys.net/parsing-email-using-python-content
# - Parse Multi-Part Email with Sub-parts using Python, http://stackoverflow.com/a/4825114/1435167

def mylistdir(directory):
    """A specialized version of os.listdir() that ignores files that
    start with a leading period."""
    filelist = os.listdir(directory)
    return [x for x in filelist
            if not (x.startswith('.'))]

def openmailbox(inmailboxpath, outmailboxpath):
    """ Open a mailbox (maildir) at the given path and cycle
    on all te given emails.
    """
    # If Factory = mailbox.MaildirMessage or rfc822.Message  any update moves the email in /new from /cur
    # see > http://stackoverflow.com/a/13445253/1435167
    mbox = mailbox.Maildir(inmailboxpath, factory=None)
    # iterate all the emails in the hierarchy
    for key, msg in mbox.iteritems():
        # ToDo Skip messages without 'attachment' without parsing parts,but what are attachments?
        #      I retain text/plain and text/html.
        # if 'alternative' in msg.get_content_type():
        # if msg.is_multipart():
        headers_as_str = u"\n".join([
            u"{}={}".format(k, ensure_unicode(msg[k]))
            for k in sorted(msg.keys())
            # avoid custom headers, prone to be mutatable
            if not k.lower().startswith("x-")
        ])
        headers_as_bytes = headers_as_str.encode("utf-8")
        outpath = outmailboxpath + hashlib.md5(headers_as_bytes).hexdigest() + "/"

        print 'Key          : ',key
        print 'Subject      : ',msg.get('Subject')
        print 'Outpath      : ',outpath
        if options.verbose:
            print 'Multip.      : ',msg.is_multipart()
            print 'Content-Type : ',msg.get('Content-Type')
            print 'Parts        : '
        detach(msg, key, outpath, mbox)
        print '='*20

def detach(msg, key, outpath, mbox):
    """ Cycle all the part of message,
    detach all the not text or multipart content type to outmailboxpath
    delete the header and rewrite is as a text inline message log.
    """
    print '-----'
    for part in msg.walk():
        content_maintype = part.get_content_maintype()
        if content_maintype == "multipart":
            continue
        if content_maintype == "text":
            # only small texts must be kept. Others are most likely disguised
            # attachments
            if len(part.get_payload()) < 500 * 1000:
                continue
        if part.get_content_type().startswith("application/pgp-"):
            # signatures are not worth consuming a separated file
            continue
        if part.get_content_type() == "multipart/signed":
            # signatures are not worth consuming a separated file
            continue
        filename = part.get_filename()
        if options.verbose:
            print '   Content-Disposition  : ', part.get('Content-Disposition')
            print '   maintytpe            : ',part.get_content_maintype()
        print '    %s : %s' % (part.get_content_type(), filename)
        try:
            os.makedirs(outpath)
        except OSError:
            if not os.path.isdir(outpath):
                raise
        if filename is not None:
            filename = ensure_unicode(filename)
            filename = decode(filename)
            if "?" in filename:
                filename = filename.split("?")[0]
            filename = filename\
                       .replace("\n", "")\
                       .replace("\r", "")\
                       .replace("\l", "")\
                       .replace("/", "_")
            filename = filename.encode("ascii", "xmlcharrefreplace").decode("ascii")
        else:
            fp = tempfile.NamedTemporaryFile(
                dir=outpath,
                delete=False,
                suffix="." + part.get_content_subtype(),
            )
            filename = os.path.basename(fp.name)
            print("Computed the filename {}".format(fp.name))
            fp.close()
        if options.save_attach:
            filepath = os.path.join(outpath, filename)
            if os.path.exists(filepath) or os.path.islink(filepath):
                base, ext = os.path.splitext(filename)
                fp = tempfile.NamedTemporaryFile(
                    dir=outpath,
                    prefix=base,
                    suffix=ext,
                    delete=False
                )
            else:
                fp = open(filepath, 'wb')
            fp.write(part.get_payload(decode=1) or "")
            fp.close()
        else:
            print("Not saving attachment, use -s to save them")
        outmessage = '    ATTACHMENT=%s\n    saved into\n    OUTPATH=%s' %(filename,outpath[len(OUTPATH):]+filename)
        if options.del_attach:
            # rewrite header and delete attachment in payload
            for h in part.keys():
                del part[h]
            part.set_payload(outmessage)
            part.set_param('Content-Type','text/html; charset=ISO-8859-1')
            part.set_param('Content-Disposition','inline')
            mbox[key] = msg
            outmessage += " and deleted from message"
        print outmessage
        print '-----'

# Recreate flat IMAP folder structure as directory structure
# WARNING: If foder name contains '.' it will changed to os.sep and it will creare a new subfolder!!!
for folder in mylistdir(PATH):
    folderpath = os.path.join(OUTPATH, folder.replace('.',os.sep)+os.sep)
    try:
        os.makedirs(folderpath)
    except OSError:
        if not os.path.isdir(folderpath):
            raise
    print
    print 'Opening mailbox:',PATH+os.sep+folder
    print '  Output folder: ',folderpath
    print
    print '='*20
    try:
        openmailbox(PATH+os.sep+folder, folderpath)
    except:
        import sys
        import ipdb
        ipdb.post_mortem(sys.exc_info()[2])
    print 40*'*'
