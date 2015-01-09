Public Class CS_form

    Private Sub initialize_scripts_link_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles initialize_scripts_link.LinkClicked
        System.Diagnostics.Process.Start("Q:\Blue Zone Scripts\Script Files\ACTIONS - initialize scripts.vbs")
        Application.Exit()
    End Sub

    Private Sub install_EA_notebooks_link_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles install_EA_notebooks_link.LinkClicked
        System.Diagnostics.Process.Start("Q:\OneNote\Installation scripts\open EA notebooks.vbs")
        Application.Exit()
    End Sub

    Private Sub CS_form_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub CourseMill_link_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles CourseMill_link.LinkClicked
        System.Diagnostics.Process.Start("https://elearning.co.anoka.mn.us/cm6/cm0682/home.html")
        Application.Exit()
    End Sub

    Private Sub Employee_Online_link_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles Employee_Online_link.LinkClicked
        System.Diagnostics.Process.Start("https://emponline.co.anoka.mn.us")
        Application.Exit()
    End Sub

    Private Sub Intranet_link_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles Intranet_link.LinkClicked
        System.Diagnostics.Process.Start("http://irp.co.anoka.mn.us/")
        Application.Exit()
    End Sub

    Private Sub Public_Website_link_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles Public_Website_link.LinkClicked
        System.Diagnostics.Process.Start("http://www.anokacounty.us")
        Application.Exit()
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles SIR_link.LinkClicked
        System.Diagnostics.Process.Start("https://www.dhssir.cty.dhs.state.mn.us/Pages/Default.aspx")
        Application.Exit()
    End Sub
End Class