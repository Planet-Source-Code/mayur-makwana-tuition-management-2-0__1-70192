VERSION 5.00
Begin VB.Form frmView 
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ctlCRViewer As CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer
Attribute ctlCRViewer.VB_VarHelpID = -1

Dim WithEvents ctlCommandExit As VB.CommandButton
Attribute ctlCommandExit.VB_VarHelpID = -1
Dim WithEvents ctlCommandDisplayReport As VB.CommandButton
Attribute ctlCommandDisplayReport.VB_VarHelpID = -1

Dim crxApplication As New CRAXDRT.Application
Dim crxReport As New CRAXDRT.Report



Private Sub Form_Load()
 ' Dynamically add the CRVIEWER control to the form
   Set ctlCRViewer = Controls.Add("CrystalReports.ActiveXReportViewer", "ctlCRViewer", frmView)
   
   'Make the Viewer visible
   ctlCRViewer.Visible = True
   
   ' Set the location and size of the CRViewer
   ctlCRViewer.Move 0, 0, Me.Width, Me.Height
   ctlCRViewer.Visible = True
Select Case RptName
   Case "TestType Wise"
        Set crxReport = crxApplication.OpenReport(App.Path & "\REPORTS\Test & Result\TestType Wise.RPT")
   Case "Test Date Wise"
        Set crxReport = crxApplication.OpenReport(App.Path + "\REPORTS\Test & Result\Test Date Wise.RPT")
   Case "Roll No. Wise"
        Set crxReport = crxApplication.OpenReport(App.Path + "\REPORTS\Test & Result\Roll No. Wise.RPT")
   Case "Standard Wise"
        Set crxReport = crxApplication.OpenReport(App.Path + "\REPORTS\Test & Result\Standard Wise.RPT")
   Case "Standard Wise"
        Set crxReport = crxApplication.OpenReport(App.Path + "\REPORTS\Fees Detail\Standard Wise.RPT")
   Case "Branch Wise"
        Set crxReport = crxApplication.OpenReport(App.Path + "\REPORTS\Fees Detail\Branch Wise.RPT")
   Case "Status Wise"
        Set crxReport = crxApplication.OpenReport(App.Path + "\REPORTS\Attendance\Status Wise.RPT")
   Case "Debit Fees Wise"
        Set crxReport = crxApplication.OpenReport(App.Path + "\REPORTS\Fees Detail\Debit Fees Wise.RPT")
   Case "Batch Wise"
        Set crxReport = crxApplication.OpenReport(App.Path + "\REPORTS\Attendance\Batch Wise.RPT")
   Case "Att.Date Wise"
        Set crxReport = crxApplication.OpenReport(App.Path + "\REPORTS\Attendance\Att.Date Wise.RPT")
   Case "Subject Wise"
        Set crxReport = crxApplication.OpenReport(App.Path + "\REPORTS\Test & Result\Subject Wise.RPT")
   Case "Sector Wise"
        Set crxReport = crxApplication.OpenReport(App.Path + "\REPORTS\Student\Sector Wise.RPT")
   Case "Surname Wise"
        Set crxReport = crxApplication.OpenReport(App.Path + "\REPORTS\Student\Surname Wise.RPT")
   Case "Standard Wise"
        Set crxReport = crxApplication.OpenReport(App.Path + "\REPORTS\Student\Standard Wise.RPT")
   Case "Branch Wise"
        Set crxReport = crxApplication.OpenReport(App.Path + "\REPORTS\Student\Branch Wise.RPT")
   Case "Addmission Date Wise"
        Set crxReport = crxApplication.OpenReport(App.Path + "\REPORTS\Student\Addmission Date Wise.RPT")
   Case "School Wise"
        Set crxReport = crxApplication.OpenReport(App.Path + "\REPORTS\Student\School Wise.RPT")
   Case "Result Wise"
        Set crxReport = crxApplication.OpenReport(App.Path + "\REPORTS\Test & Result\Result Wise.RPT")
End Select
   ctlCRViewer.EnableNavigationControls = True
   ctlCRViewer.EnableSearchControl = False
   ctlCRViewer.ReportSource = crxReport
   ctlCRViewer.ViewReport
End Sub
