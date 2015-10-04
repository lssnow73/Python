# -*- coding: cp949 -*-
# =============================================================================
# Title : Mail 가져오기
# =============================================================================

import sys
import locale
import win32com.client as win32
from time import sleep


ReportDayOfWeek = "09/21/15"
SubjectPrefix   = "[주간계획]".decode('cp949')
PostBox         = "sslee@ubiquoss.com"


ListOfMyTeam = []

def CreateTeamMemberListbyFile(DisplayOn = True):
    try:
        with open('MemberOfTeam.txt', 'r') as  MemberOfTeam_file:
            for line in MemberOfTeam_file:
                if(DisplayOn):
                    print line
                ListOfMyTeam.append(line.decode('cp949'))

    except IOError as err:
        print('File error: ' + str(err))

    if(DisplayOn):
        for list in ListOfMyTeam:
            print list


def DisplayOfficeOutlookItem(msg):
    print 'Subject: ' + msg.Subject
    print 'SenderName: ' + msg.SenderName
    print 'SenderEmailAddress: ' + msg.SenderEmailAddress
    print 'To: ' + msg.To
    print 'CC: ' + msg.CC
    print 'ReceivedByName: ' + msg.ReceivedByName
    print 'ReceivedTime: ' + str(msg.ReceivedTime)
    print 'Size: ' + str(msg.Size)
    #print unicode(msg.Body)
    return


AllMemberOutlookMessage = {}

def OfficeOutlook(PersonalFolderName = "주간업무관련", AppClose = False):
    app = win32.Dispatch("Outlook.Application")
    NameSpace = app.GetNamespace("MAPI")
    PersonalFolder = NameSpace.Folders.Item(PostBox)

    Inbox = PersonalFolder.Folders.Item(PersonalFolderName)

    for i in range(len(Inbox.Items), len(Inbox.Items)-12, -1):
        msg = Inbox.Items.Item(i)
        (ReceivedDate, ReceivedTime) = str(msg.ReceivedTime).split(' ', 1)
        if ReceivedDate != ReportDayOfWeek:
            continue

        if SubjectPrefix not in msg.Subject:
            continue
        if msg.Subject.startswith(SubjectPrefix):
            DisplayOfficeOutlookItem(msg)
        else:
            print "Do Not Start Prefix"

    if(AppClose):
        app.Application.Quit()

    return msg.Body


def OfficeWord(MSGBody, AppClose = False):
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = 1
    sleep(1)
    doc = word.Documents.Open("d:\\tmp\\aaa.docx")
    tbl = doc.Tables.Item(1)
    sleep(1)

    tbl.Cell(2,2).Range.Delete(1, 1)
    sleep(2)
    tbl.Cell(2,2).Range.InsertAfter(MSGBody)
    sleep(2)

    doc.Close()
    sleep(1)
    word.Quit()


if __name__ == '__main__':
    CreateTeamMemberListbyFile(DisplayOn=False)
#    MessageBody = OfficeOutlook()
#    OfficeWord(MessageBody)

# End of File