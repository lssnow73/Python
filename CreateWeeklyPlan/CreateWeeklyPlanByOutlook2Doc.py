# -*- coding: cp949 -*-
# =============================================================================
# Title : Mail 가져오기
# =============================================================================

import sys
import locale
import win32com.client as win32
from time import sleep
from __builtin__ import True


ReportDayOfWeek = "01/04/16"
SubjectPrefix   = "[주간계획세부사항]".decode('cp949')
DefaultPostBox  = "sslee@ubiquoss.com"
DefaultDisplay  = True
DefaultFolder   = "주간계획"
DefaultOpenFile = "D:\\회사업무_기타\\주간업무수행_2016\\SW2G_주간업무수행 양식-01주차_160104-SW5.docx"


AllTeamMember = {}

def CreateTeamMemberListbyFile(DisplayOn = DefaultDisplay):
    number = 0
    
    try:
        with open('MemberOfTeam.txt', 'r') as  MemberOfTeam_file:
            for line in MemberOfTeam_file:
                (name, rank) = line.decode('cp949').split()
                number = number + 1
                AllTeamMember['{:0=2}'.format(number)] = (name, rank)

    except IOError as err:
        print('File error: ' + str(err))

    if(DisplayOn):
        for TeamMember in sorted(AllTeamMember.keys()):
            print TeamMember + AllTeamMember[TeamMember][0] + AllTeamMember[TeamMember][1]
    return


def DisplayItem_OfficeOutlook(msg):
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


def OpenMailBox_OfficeOutlook(PostBox = DefaultPostBox, PersonalFolderName = DefaultFolder):
    global OfficeOutlook
    OfficeOutlook = win32.Dispatch("Outlook.Application")
    NameSpace = OfficeOutlook.GetNamespace("MAPI")
    PersonalFolder = NameSpace.Folders.Item(PostBox)
    print PersonalFolder.Folders
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
                DisplayItem_OfficeOutlook(msg)
            return msg
    return


def OpenDocx_OfficeWord(OpenFile = DefaultOpenFile):
    global OfficeWord
    OfficeWord = win32.gencache.EnsureDispatch('Word.Application')
    OfficeWord.Visible = 1
    OpenDocx = OfficeWord.Documents.Open(OpenFile)
    sleep(1)

    return OpenDocx


def ChangeMailMessageBody_OfficeWord(MailMsgBody):
    if MailMsgBody == None:
        return None
    
    FormattedBody = MailMsgBody.replace(u'굕a1', u'굑f0a1')
    FormattedBody = MailMsgBody.replace(u'굕a0', ' ')
    FormattedBody = FormattedBody.replace(u'굕f0', u'굑f0f0')

    return FormattedBody


def WriteDocx_OfficeWord(OpenDocx, TeamMember, MailMsg, OpenFile = DefaultOpenFile):
    global AllTeamMember

    if MailMsg == None:
        return
    
    index = int(TeamMember) + 1
    MemberName = AllTeamMember[TeamMember][0] + " " + AllTeamMember[TeamMember][1]
    
    if OpenDocx == None:
        OpenDocx = OpenDocx_OfficeWord(OpenFile)

    tbl = OpenDocx.Tables.Item(1)

    tbl.Cell(index, 1).Range.Delete(1, 1)
    tbl.Cell(index, 2).Range.Delete(1, 1)
    sleep(2)
    tbl.Cell(index, 1).Range.InsertAfter(MemberName)
    sleep(2)
    FormattedBody = ChangeMailMessageBody_OfficeWord(MailMsg.Body)
    tbl.Cell(index, 2).Range.InsertAfter(FormattedBody)
    #sleep(2)
    """
    tbl.Cell(index, 2).Range.Font.Name = "맑은 고딕"
    tbl.Cell(index, 2).Range.Font.Bold = 1
    tbl.Cell(index, 2).Range.Font.Size = 13
    tbl.Cell(index, 2).Range.Characters.Count
    tbl.Cell(index, 2).Range.FormattedText = FormattedBody
    sleep(2)
    #    tbl.Cell(index, 2).Range.Find.Execute('U3204')
    #    print unicode(tbl.Cell(index, 2).Range.Text[0:1])
    """
    return


if __name__ == '__main__':
    CreateTeamMemberListbyFile()
    MailBox = OpenMailBox_OfficeOutlook()
    Docx    = OpenDocx_OfficeWord()

    for TeamMember in sorted(AllTeamMember.keys()):
        MailMessage = SearchMailBox_OfficeOutlook(MailBox, AllTeamMember[TeamMember][0])
        if(MailMessage):
            DisplayItem_OfficeOutlook(MailMessage)
            WriteDocx_OfficeWord(Docx, TeamMember, MailMessage)

    """
    TeamMember = "01"
    MailMessage = SearchMailBox_OfficeOutlook(MailBox, AllTeamMember[TeamMember][0])
    if(MailMessage):
        WriteDocx_OfficeWord(Docx, TeamMember, MailMessage)
    """

    Docx.Close()
    sleep(1)
    OfficeWord.Quit()
#    OfficeOutlook.Application.Quit()

# End of File