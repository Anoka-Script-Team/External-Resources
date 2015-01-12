Public Class PA_form

    Private Property list_path As String




    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim user_ID
        user_ID = SystemInformation.UserName()
        list_path = "Q:\Blue Zone Scripts\Spreadsheets for script use\worker list.xlsx"
        If System.IO.File.Exists(list_path) = True Then
            Dim objExcel As Object
            objExcel = CreateObject("Excel.Application")
            objExcel.Visible = False
            Dim objWorkbook As Object
            objWorkbook = objExcel.Workbooks.Open(list_path)
            Dim worker_unit
            Dim excel_row
            excel_row = 2
            Do
                If objExcel.cells(excel_row, 3).Value = UCase(user_ID) Then
                    worker_unit = objExcel.cells(excel_row, 4).Value
                    If worker_unit = "CS1" Or worker_unit = "CS2" Or worker_unit = "CS3" Or worker_unit = "CS4" Then
                        My.Forms.CS_form.Show()
                        Me.Close()
                    End If
                    Exit Do
                End If
                excel_row = excel_row + 1
            Loop Until objExcel.cells(excel_row, 4).Value = ""
            objExcel.Quit()
        Else
            GroupBox5.Hide()
        End If
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles Combined_Manual_link.LinkClicked
        System.Diagnostics.Process.Start("http://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CombinedManual")
        Application.Exit()
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles HCPM_link.LinkClicked
        System.Diagnostics.Process.Start("http://hcopub.dhs.state.mn.us/hcpmstd/")
        Application.Exit()
    End Sub

    Private Sub LinkLabel5_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles VerifyMN_link.LinkClicked
        System.Diagnostics.Process.Start("https://smi.dhs.state.mn.us/login.do")
        Application.Exit()
    End Sub

    Private Sub LinkLabel14_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles Case_transfer_email_link.LinkClicked
        System.Diagnostics.Process.Start("L:\Email Templates\Case Transfer Request.oft")
        Application.Exit()
    End Sub

    Private Sub LinkLabel10_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles Employee_Online_link.LinkClicked
        System.Diagnostics.Process.Start("https://emponline.co.anoka.mn.us")
        Application.Exit()
    End Sub

    Private Sub LinkLabel13_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles Property_records_link.LinkClicked
        System.Diagnostics.Process.Start("https://prtinfo.co.anoka.mn.us/(oazyfc4522vzofmkapfc2byx)/search.aspx")
        Application.Exit()
    End Sub

    Private Sub LinkLabel12_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles HS_Apps_link.LinkClicked
        'Grabbing user ID for msgbox
        Dim objNet
        Dim user_ID
        objNet = CreateObject("WScript.NetWork")
        user_ID = UCase(objNet.UserName)

        'Opening an IE window, navigating to HS-Apps and entering the user name
        Dim IE ' As Object
        IE = CreateObject("InternetExplorer.Application")
        IE.visible = True
        IE.Navigate("http://hsapps.co.anoka.mn.us/")
        Do
            Threading.Thread.Sleep(100)
        Loop Until IE.Busy = False
        IE.Document.All.Item("ctl00$cphBody01$txtUserName").Value = user_ID
        IE.Document.Forms(0).Submit()

        'Exiting application
        Application.Exit()
    End Sub

    Private Sub LinkLabel11_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles CourseMill_link.LinkClicked
        System.Diagnostics.Process.Start("https://elearning.co.anoka.mn.us/cm6/cm0682/home.html")
        Application.Exit()
    End Sub

    Private Sub LinkLabel15_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles ACCAP_resource_guide_link.LinkClicked
        System.Diagnostics.Process.Start("http://www.accap.org/Documents/Resource_Guide/AnokaCountyResourceGuideVol-13-2014-2016.pdf")
        Application.Exit()
    End Sub

    Private Sub LinkLabel17_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles Intranet_link.LinkClicked
        System.Diagnostics.Process.Start("http://irp.co.anoka.mn.us/")
        Application.Exit()
    End Sub

    Private Sub GroupBox4_Enter(sender As Object, e As EventArgs) Handles GroupBox4.Enter
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Application.Exit()
    End Sub

    Private Sub LinkLabel6_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles eDocs_link.LinkClicked
        System.Diagnostics.Process.Start("http://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&Redirected=true&dDocName=id_000100")
        Application.Exit()
    End Sub

    Private Sub LinkLabel3_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles MCRE_income_worksheet_link.LinkClicked
        System.Diagnostics.Process.Start("https://edocs.dhs.state.mn.us/lfserver/public/DHS-3352-ENG")
        Application.Exit()
    End Sub

    Private Sub LinkLabel4_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles MMIS_user_manual_link.LinkClicked
        System.Diagnostics.Process.Start("http://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=user_manual")
        Application.Exit()
    End Sub

    Private Sub LinkLabel18_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles Trainlink_link.LinkClicked
        System.Diagnostics.Process.Start("http://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=Training")
        Application.Exit()
    End Sub

    Private Sub LinkLabel16_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles Public_Website_link.LinkClicked
        System.Diagnostics.Process.Start("http://www.anokacounty.us")
        Application.Exit()
    End Sub

    Private Sub LinkLabel7_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles IAPM_manual_link.LinkClicked
        System.Diagnostics.Process.Start("http://hcopub.dhs.state.mn.us/iapmstd/")
        Application.Exit()
    End Sub

    Private Sub LinkLabel8_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles MNsure_link.LinkClicked
        Shell("C:\Program Files (x86)\Mozilla Firefox\firefox.exe https://www.mnsure.org/")
        Application.Exit()
    End Sub

    Private Sub LinkLabel9_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles OneSource_link.LinkClicked
        System.Diagnostics.Process.Start("http://hcopub.dhs.state.mn.us/onesourcestd")
        Application.Exit()
    End Sub

    Private Sub LinkLabel19_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles NADA_link.LinkClicked
        System.Diagnostics.Process.Start("http://www.nadaguides.com/Cars/Research-Center")
        Application.Exit()
    End Sub

    Private Sub LinkLabel20_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles SAVE_link.LinkClicked
        System.Diagnostics.Process.Start("https://save.uscis.gov/Web/vislogin.aspx?JS=YES")
        Application.Exit()
    End Sub

    Private Sub LinkLabel21_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles Work_Number_link.LinkClicked
        System.Diagnostics.Process.Start("http://www.theworknumber.com/SocialServices/")
        Application.Exit()
    End Sub

    Private Sub LinkLabel22_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles Xcel_Energy_link.LinkClicked
        System.Diagnostics.Process.Start("https://eap.xcelenergy.com/portal/site/eaportal/template.LOGIN/")
        Application.Exit()
    End Sub


    

    Private Sub Starlite_link_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles Starlite_link.LinkClicked
        System.Diagnostics.Process.Start("https://starlite.co.anoka.mn.us/(ontbzi21bzksbw453txmebmu)/login.aspx")
        Application.Exit()
    End Sub

    Private Sub SIR_link_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)

    End Sub

    Private Sub install_EA_notebooks_link_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles install_EA_notebooks_link.LinkClicked
        System.Diagnostics.Process.Start("Q:\OneNote\Installation scripts\open EA notebooks.vbs")
        Application.Exit()
    End Sub

    Private Sub SIR_link_LinkClicked_1(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles SIR_link.LinkClicked
        System.Diagnostics.Process.Start("https://www.dhssir.cty.dhs.state.mn.us/Pages/Default.aspx")
        Application.Exit()
    End Sub

    Private Sub initialize_scripts_link_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles initialize_scripts_link.LinkClicked
        System.Diagnostics.Process.Start("Q:\Blue Zone Scripts\Script Files\ACTIONS - initialize scripts.vbs")
        Application.Exit()
    End Sub

End Class

