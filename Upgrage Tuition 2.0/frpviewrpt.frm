VERSION 5.00
Object = "{3C62B3DD-12BE-4941-A787-EA25415DCD27}#10.0#0"; "crviewer.dll"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Dynamically Add Controls.vbp"
   ClientHeight    =   9795
   ClientLeft      =   2460
   ClientTop       =   525
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   ScaleHeight     =   9795
   ScaleWidth      =   10050
   Begin CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer CrystalActiveXReportViewer1 
      Height          =   1620
      Left            =   2370
      TabIndex        =   0
      Top             =   4320
      Width           =   2730
      lastProp        =   600
      _cx             =   4815
      _cy             =   2857
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************
'Author:  Pedro Gomes
'Created:  January 2001
'Modified: March 25th by Trista (updated for version 10)
'***********************

'This Microsoft Visual Basic application uses the Report Designer Component (RDC)
'Automation Server (Craxdrt.dll) as the reporting development tool.

'This sample application adds two command buttons and the Crystal Report
'Viewer control to a form at runtime.  In addition, this
'application demonstrates how to program the events of a dynamically added control.

'This Project requires that the following are referenced: (under the Project | References menu)
' 1. Crystal Report ActiveX Designer Runtime Library
' 2. Crystal Report Viewer Control

'IMPORTANT:
'Before running this application, go to the Project | Properties menu.
'Under the 'Make' tab, uncheck "Remove information about unused ActiveX controls".


'*************************************************************************************
'General Declarations of Form1
'Once we have our variables declared we can start building our Application and
'Report object as well as . The Application object is the object we will use to open
'and set our Report object. The Report object is used to get and set report properties
'at runtime. Under the Form_Load event of Form1 place the following lines of
'code:
'*************************************************************************************

Dim WithEvents ctlCRViewer As CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer
Attribute ctlCRViewer.VB_VarHelpID = -1

Dim WithEvents ctlCommandExit As VB.CommandButton
Attribute ctlCommandExit.VB_VarHelpID = -1
Dim WithEvents ctlCommandDisplayReport As VB.CommandButton
Attribute ctlCommandDisplayReport.VB_VarHelpID = -1

Dim crxApplication As New CRAXDRT.Application
Dim crxReport As New CRAXDRT.Report

'-------------------------------------------------------------------------------------


Private Sub Form_Load()

   ' Dynamically add a command button control to the form (CommandButton1)
   Set ctlCommandExit = Controls.Add("VB.CommandButton", _
                    "ctlCommand1", Form1)
  
   ' Set the location and size of CommandButton1.
   ctlCommandExit.Move 6600, 550, 1335, 400
   
     
   ' Set the caption
   ctlCommandExit.Caption = "Exit"
   

   ' Make it visible
   ctlCommandExit.Visible = True
   
   
   
   ' Dynamically add a command button control to the form (CommandButton2)
   Set ctlCommandDisplayReport = Controls.Add("VB.CommandButton", _
                    "ctlCommand2", Form1)
  
   ' Set the location and size of CommandButton2.
   ctlCommandDisplayReport.Move 6600, 75, 1335, 400
   
     
   ' Set the caption
   ctlCommandDisplayReport.Caption = "Display Report"
   
   ' Make it visible
   ctlCommandDisplayReport.Visible = True
   
   'Set focus to this control when form loads
   ctlCommandDisplayReport.TabIndex = 0
   
   ' Dynamically add the CRVIEWER control to the form
   Set ctlCRViewer = Controls.Add("CrystalReports.ActiveXReportViewer", "ctlCRViewer", Form1)
   
   'Make the Viewer visible
   ctlCRViewer.Visible = True
   
   ' Set the location and size of the CRViewer
   ctlCRViewer.Move 0, 1000, 10000, 8800

    
End Sub

Private Sub ctlCommandExit_Click()
   'Unloads the form and ends the program
   Unload Me

End Sub

Private Sub ctlCommandDisplayReport_Click()
   ctlCRViewer.Visible = True
   Set crxReport = crxApplication.OpenReport(App.Path & "\REPORTS\Test & Result\TestType Wise.RPT")
   ctlCRViewer.EnableNavigationControls = True
   ctlCRViewer.EnableSearchControl = False
   ctlCRViewer.ReportSource = crxReport
   ctlCRViewer.ViewReport
'   ' Make the CRViewer visible
'   ctlCRViewer.Visible = True
'
'   'Set the report object
'   Set crxReport = crxApplication.OpenReport(App.Path & "\chart.rpt")
'
'   'Enables the CRViewer's Navigation Controls
'   ctlCRViewer.EnableNavigationControls = True
'
'   'Disables the CRViewer's Search control
'   ctlCRViewer.EnableSearchControl = False
'
'   'maximizes the window
'   'Me.WindowState = 2
'
'
'
'   'Another way to define a position would be:
'   'ctlCRViewer.Top = 1000
'   'ctlCRViewer.Left = 0
'   'ctlCRViewer.Width = ScaleWidth
'   'ctlCRViewer.Height = ScaleHeight
'
'   'Set the report source
'   ctlCRViewer.ReportSource = crxReport
'
'   'View the report
'   ctlCRViewer.ViewReport

   
End Sub


Private Sub Form_Unload(Cancel As Integer)
 'Destroy the report object
 Set crxReport = Nothing
End Sub
