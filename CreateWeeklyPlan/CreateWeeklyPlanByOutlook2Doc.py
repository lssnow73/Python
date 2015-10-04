# -*- coding: cp949 -*-
# =============================================================================
# Title : Mail 가져오기
# =============================================================================

import sys
import locale
import win32com.client as win32
from time import sleep
from __builtin__ import True


ReportDayOfWeek = "09/21/15"
SubjectPrefix   = "[주간계획]".decode('cp949')
DefaultPostBox  = "sslee@ubiquoss.com"
DefaultDisplay  = False
DefaultOpenFile = "d:\\tmp\\aaa.docx"


AllTeamMember = {}

def CreateTeamMemberListbyFile(DisplayOn = DefaultDisplay):
    try:
        with open('MemberOfTeam.txt', 'r') as  MemberOfTeam_file:
            for line in MemberOfTeam_file:
                (number, name, rank) = line.decode('cp949').split()
                AllTeamMember[number] = (name, rank)

    except IOError as err:
        print('File error: ' + str(err))

    if(DisplayOn):
        for TeamMember in sorted(AllTeamMember.keys()):
            print TeamMember + AllTeamMember[TeamMember][0] + AllTeamMember[TeamMember][1]
    return


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


def OpenMailBox_OfficeOutlook(PostBox = DefaultPostBox, PersonalFolderName = "주간업무관련"):
    global OfficeOutlook
    OfficeOutlook = win32.Dispatch("Outlook.Application")
    NameSpace = OfficeOutlook.GetNamespace("MAPI")
    PersonalFolder = NameSpace.Folders.Item(PostBox)
    Inbox = PersonalFolder.Folders.Item(PersonalFolderName)

    return Inbox


def SearchMailBox_OfficeOutlook(Inbox, SearchUser, SearchDate = ReportDayOfWeek,
                                SearchSubject = SubjectPrefix, DisplayOn = DefaultDisplay):
    if Inbox == None:
        return False

    for i in range(len(Inbox.Items), len(Inbox.Items)-50, -1):
        msg = Inbox.Items.Item(i)
        (ReceivedDate, ReceivedTime) = str(msg.ReceivedTime).split(' ', 1)
        if ReceivedDate != SearchDate:
            continue
        if SearchSubject not in msg.Subject:
            continue
        if not msg.Subject.startswith(SearchSubject):
            continue

        if SearchUser in msg.Subject:
            if(DisplayOn):
                DisplayOfficeOutlookItem(msg)
            return msg
    return


def OpenDocx_OfficeWord(OpenFile = DefaultOpenFile):
    global OfficeWord
    OfficeWord = win32.gencache.EnsureDispatch('Word.Application')
    OfficeWord.Visible = 1
    OpenDocx = OfficeWord.Documents.Open(OpenFile)
    sleep(1)

    return OpenDocx

    
def WriteDocx_OfficeWord(OpenDocx, TeamMember, MailMsg, OpenFile = DefaultOpenFile):
    global AllTeamMember

    if MailMsg == None:
        return
    
    index = int(TeamMember) + 1
    MemberName = AllTeamMember[TeamMember][0] + " " + AllTeamMember[TeamMember][1]
    
    if OpenDocx == None:
        OpenDocx = OpenDocx_OfficeWord(OpenFile)

    tbl = OpenDocx.Tables.Item(1)
    sleep(1)

    tbl.Cell(index, 1).Range.Delete(1, 1)
    tbl.Cell(index, 2).Range.Delete(1, 1)
    sleep(1)
    tbl.Cell(index, 1).Range.InsertAfter(MemberName)
    tbl.Cell(index, 2).Range.InsertAfter(MailMsg.Body)
    sleep(1)
    return


if __name__ == '__main__':
    CreateTeamMemberListbyFile()
    MailBox = OpenMailBox_OfficeOutlook()
    Docx    = OpenDocx_OfficeWord()
    
    for TeamMember in sorted(AllTeamMember.keys()):
        MailMessage = SearchMailBox_OfficeOutlook(MailBox, AllTeamMember[TeamMember][0])
        if(MailMessage):
            DisplayOfficeOutlookItem(MailMessage)
            WriteDocx_OfficeWord(Docx, TeamMember, MailMessage)

    Docx.Close()
    sleep(1)
    OfficeWord.Quit()
#    OfficeOutlook.Application.Quit()

# End of File