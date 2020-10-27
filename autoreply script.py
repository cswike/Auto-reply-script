#
# Find last mail from specific sender in inbox
# Open the mail with reply all function
#

import win32com.client as win32
import os


# nice information to work with windows
# excel and so on
# https://pbpython.com/windows-com.html
# Complete documentation here: https://docs.microsoft.com/en-us/office/vba/api/outlook.mailitem.move

outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Here the name of the folders
# 3 = Drafts
# 4 = Outbox
# 5 = Sent items
# 6 = Inbox
inboxItems = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case,
                                         # the inbox. You can change that number to reference
                                         # any other folder

messages_recd = inboxItems.Items.restrict("@SQL=%today(""urn:schemas:httpmail:datereceived"")%")

found = False
message_sender_to_find = 'sender_here@email.com'
sender_found = ''
check_subj = ''

for message in messages_recd: 
    # print('Message class: {}'.format(message.Class))
    # print('Message sender: {}'.format(message.SenderEmailAddress.encode("utf-8")))
    # Message class 43 -> mail
    if message.Class == 43: 
        if message_sender_to_find in message.SenderEmailAddress:
                sender_found = message.SenderEmailAddress
                found = True
                while True:
                    try:
                        check_subj = input("Message to reply to has subject: " + message.Subject + ". Reply to this one? Y/N: ")
                        if check_subj not in {'y','Y','n','N'}:
                            print("Sorry, I didn't understand that.")
                            continue
                        else:
                            break
                    except ValueError:
                        print("Sorry, I didn't understand that.")
                        continue
                if check_subj in {'y','Y'}:
                    replyAll = message.ReplyAll()
                    replyAll.Body = "Email Body Here"
                    replyAll.Send()
                    break
                elif check_subj in {'n','N'}:
                    found = False
                    continue

if found:
  print('Done for this item !! -> {}'.format(sender_found))
else:
  print('Message not found')

os.system('pause')
