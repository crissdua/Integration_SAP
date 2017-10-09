Imports System.Windows.Forms
Imports System.IO
Imports System.Xml
Imports System.Threading
Public Class Form1
    Private Property _oCompany As SAPbobsCOM.Company

    Public Property oCompany() As SAPbobsCOM.Company
        Get
            Return _oCompany
        End Get
        Set(ByVal value As SAPbobsCOM.Company)
            _oCompany = value
        End Set
    End Property

    Public Function MakeConnectionSAP() As Boolean
        Dim Connected As Boolean = False
        '' Dim cnnParam As New Settings
        Try
            Connected = -1

            oCompany = New SAPbobsCOM.Company
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014
            oCompany.DbUserName = "sa"
            oCompany.DbPassword = "12345"
            oCompany.Server = "DESKTOP-13FMJTF"
            oCompany.CompanyDB = "FYA"
            oCompany.UserName = "manager"
            oCompany.Password = "alegria"
            oCompany.LicenseServer = "DESKTOP-13FMJTF:30000"
            oCompany.UseTrusted = False
            Connected = oCompany.Connect()

            If Connected <> 0 Then
                Connected = False
                MsgBox(oCompany.GetLastErrorDescription)
            Else
                'MsgBox("Conexión con SAP exitosa")
                Connected = True
            End If
            Return Connected
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Dim MyThread As Thread
    Dim myStream As Stream
    'items
    Dim ItemsStart As New ThreadStart(AddressOf BackgroundItems)
    Dim CallItems As New MethodInvoker(AddressOf Me.ItemToma)
    'clientes
    Dim ClientesStart As New ThreadStart(AddressOf BackgroundClientes)
    Dim CallClientes As New MethodInvoker(AddressOf Me.ClienteToma)
    'facturas
    Dim FacturasStart As New ThreadStart(AddressOf BackgroundFacturas)
    Dim CallFacturas As New MethodInvoker(AddressOf Me.FacturasToma)
    'ordenes
    Dim OrdenStart As New ThreadStart(AddressOf BackgroundOrdenes)
    Dim CallOrdenes As New MethodInvoker(AddressOf Me.OrdenesToma)
    'traslados
    Dim TrasladoStart As New ThreadStart(AddressOf BackgroundTraslados)
    Dim CallTraslados As New MethodInvoker(AddressOf Me.TrasladosToma)
    'Nota de Credito
    Dim NcreditoStart As New ThreadStart(AddressOf BackgroundNcredito)
    Dim CallNcredito As New MethodInvoker(AddressOf Me.NcreditoToma)
    'Paciente
    Dim PacienteStart As New ThreadStart(AddressOf BackgroundPaciente)
    Dim CallPaciente As New MethodInvoker(AddressOf Me.PacienteToma)
    'Pagos_Cheque
    Dim PChequeStart As New ThreadStart(AddressOf BackgroundPCheque)
    Dim CallPCheque As New MethodInvoker(AddressOf Me.PChequeToma)
    'Pagos_Credito
    Dim PCreditoStart As New ThreadStart(AddressOf BackgroundPCredito)
    Dim CallPCredito As New MethodInvoker(AddressOf Me.PCreditoToma)
    'Pagos_Efectivo
    Dim PEfectivoStart As New ThreadStart(AddressOf BackgroundPEfectivo)
    Dim CallPEfectivo As New MethodInvoker(AddressOf Me.PEfectivoToma)
    'Pagos_Transfer
    Dim PTransferStart As New ThreadStart(AddressOf BackgroundPTransfer)
    Dim CallPTransfer As New MethodInvoker(AddressOf Me.PTransferToma)
    'cfacturas
    Dim cFacturasStart As New ThreadStart(AddressOf BackgroundcFacturas)
    Dim CallcFacturas As New MethodInvoker(AddressOf Me.cFacturasToma)
    'Pagos
    Dim PagosStart As New ThreadStart(AddressOf BackgroundPagos)
    Dim CallPagos As New MethodInvoker(AddressOf Me.PagosToma)

    Public Sub BackgroundItems()
        While True
            Me.BeginInvoke(CallItems)
            Thread.Sleep(1500)
        End While
    End Sub
    Public Sub BackgroundClientes()
        While True
            Me.BeginInvoke(CallClientes)
            Thread.Sleep(1500)
        End While
    End Sub
    Public Sub BackgroundFacturas()
        While True
            Me.BeginInvoke(CallFacturas)
            Thread.Sleep(1500)
        End While
    End Sub
    Public Sub BackgroundcFacturas()
        While True
            Me.BeginInvoke(CallcFacturas)
            Thread.Sleep(1500)
        End While
    End Sub
    Public Sub BackgroundOrdenes()
        While True
            Me.BeginInvoke(CallOrdenes)
            Thread.Sleep(1500)
        End While
    End Sub
    Public Sub BackgroundTraslados()
        While True
            Me.BeginInvoke(CallTraslados)
            Thread.Sleep(1500)
        End While
    End Sub
    Public Sub BackgroundNcredito()
        While True
            Me.BeginInvoke(CallNcredito)
            Thread.Sleep(1500)
        End While
    End Sub
    Public Sub BackgroundPaciente()
        While True
            Me.BeginInvoke(CallPaciente)
            Thread.Sleep(1500)
        End While
    End Sub
    Public Sub BackgroundPCheque()
        While True
            Me.BeginInvoke(CallPCheque)
            Thread.Sleep(1500)
        End While
    End Sub
    Public Sub BackgroundPCredito()
        While True
            Me.BeginInvoke(CallPCredito)
            Thread.Sleep(1500)
        End While
    End Sub
    Public Sub BackgroundPEfectivo()
        While True
            Me.BeginInvoke(CallPEfectivo)
            Thread.Sleep(1500)
        End While
    End Sub
    Public Sub BackgroundPTransfer()
        While True
            Me.BeginInvoke(CallPTransfer)
            Thread.Sleep(1500)
        End While
    End Sub
    Public Sub BackgroundPagos()
        While True
            Me.BeginInvoke(CallPagos)
            Thread.Sleep(1500)
        End While
    End Sub




    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Me.Opacity = 0.0R
        Me.ShowInTaskbar = False
        Me.Visible = False
        'Thread ITEMS
        MyThread = New Thread(ItemsStart)
        MyThread.IsBackground = True
        MyThread.Name = "MyThreadItems"
        MyThread.Start()
        'Thread CLIENTES
        MyThread = New Thread(ClientesStart)
        MyThread.IsBackground = True
        MyThread.Name = "MyThreadClientes"
        MyThread.Start()
        'Thread FACTURAS
        MyThread = New Thread(FacturasStart)
        MyThread.IsBackground = True
        MyThread.Name = "MyThreadFacturas"
        MyThread.Start()
        'Thread ORDENES
        MyThread = New Thread(OrdenStart)
        MyThread.IsBackground = True
        MyThread.Name = "MyThreadOrden"
        MyThread.Start()
        'Thread traslados
        MyThread = New Thread(TrasladoStart)
        MyThread.IsBackground = True
        MyThread.Name = "MyThreadTraslado"
        MyThread.Start()
        'Thread Ncredito
        MyThread = New Thread(NcreditoStart)
        MyThread.IsBackground = True
        MyThread.Name = "MyThreadNcredito"
        MyThread.Start()
        'Thread Paciente
        MyThread = New Thread(PacienteStart)
        MyThread.IsBackground = True
        MyThread.Name = "MyThreadPaciente"
        MyThread.Start()
        'Thread PCheque
        MyThread = New Thread(PChequeStart)
        MyThread.IsBackground = True
        MyThread.Name = "MyThreadPCheque"
        MyThread.Start()
        'Thread PCredito PEfectivo
        MyThread = New Thread(PCreditoStart)
        MyThread.IsBackground = True
        MyThread.Name = "MyThreadPCredito"
        MyThread.Start()
        'Thread PEfectivo
        MyThread = New Thread(PEfectivoStart)
        MyThread.IsBackground = True
        MyThread.Name = "MyThreadPEfectivo"
        MyThread.Start()
        'Thread PTransfer
        MyThread = New Thread(PTransferStart)
        MyThread.IsBackground = True
        MyThread.Name = "MyThreadPTransfer"
        MyThread.Start()
        'Thread cFACTURAS
        MyThread = New Thread(cFacturasStart)
        MyThread.IsBackground = True
        MyThread.Name = "MyThreadcFacturas"
        MyThread.Start()
        'Thread Pagos
        MyThread = New Thread(PagosStart)
        MyThread.IsBackground = True
        MyThread.Name = "MyThreadcPagos"
        MyThread.Start()
    End Sub

    Public Sub ItemToma()
        Dim objFSO As Object = CreateObject("Scripting.FileSystemObject")
        Dim objSubFolder As Object = "C:\TS\Items\"
        Dim objFolder As Object = objFSO.GetFolder(objSubFolder)
        Dim colFiles As Object = objFolder.Files

        For Each objFile In colFiles

            If existeItem() = 0 Then
                Dim entra As String = "C:\TS\Items\" + objFile.Name.ToString
                Dim integ As String = "C:\TS\Items\Integration\Item.xml"
                Dim temp As String = "C:\TS\Items\temp\Item" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml"
                File.Copy(entra, integ)
                File.Move(entra, temp)
            ElseIf existeCliente() = 1 Then
                Exit Sub
            End If
        Next
    End Sub
    Public Sub ClienteToma()
        Dim objFSO As Object = CreateObject("Scripting.FileSystemObject")
        Dim objSubFolder As Object = "C:\TS\Cliente\"
        Dim objFolder As Object = objFSO.GetFolder(objSubFolder)
        Dim colFiles As Object = objFolder.Files

        For Each objFile In colFiles

            If existeCliente() = 0 Then
                Dim entra As String = "C:\TS\Cliente\" + objFile.Name.ToString
                Dim integ As String = "C:\TS\Cliente\Integration\Cliente.xml"
                Dim temp As String = "C:\TS\Cliente\temp\Cliente" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml"
                File.Copy(entra, integ)
                File.Move(entra, temp)
            ElseIf existeCliente() = 1 Then
                Exit Sub
            End If
        Next
    End Sub
    Public Sub FacturasToma()
        Dim objFSO As Object = CreateObject("Scripting.FileSystemObject")
        Dim objSubFolder As Object = "C:\TS\Factura\"
        Dim objFolder As Object = objFSO.GetFolder(objSubFolder)
        Dim colFiles As Object = objFolder.Files

        For Each objFile In colFiles

            If existeFactura() = 0 Then
                Dim entra As String = "C:\TS\Factura\" + objFile.Name.ToString
                Dim integ As String = "C:\TS\Factura\Integration\Invoice.xml"
                Dim temp As String = "C:\TS\Factura\temp\Invoice" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml"
                File.Copy(entra, integ)
                File.Move(entra, temp)
            ElseIf existeFactura() = 1 Then
                Exit Sub
            End If
        Next
    End Sub
    Public Sub cFacturasToma()
        Try


            Dim objFSO As Object = CreateObject("Scripting.FileSystemObject")
            Dim objSubFolder As Object = "C:\TS\FacturaUpdate\"
            Dim objFolder As Object = objFSO.GetFolder(objSubFolder)
            Dim colFiles As Object = objFolder.Files
            For Each objFile In colFiles

                If existecFactura() = 0 Then
                    Dim entra As String = "C:\TS\FacturaUpdate\" + objFile.Name.ToString
                    Dim integ As String = "C:\TS\FacturaUpdate\Integration\Invoice.xml"
                    File.Move(entra, integ)
                    CancelaFactura()
                    'File.Delete("C:\TS\FacturaUpdate\Integration\Invoice.xml")
                    'File.Move(integ, temp)
                ElseIf existecFactura() = 1 Then
                    Dim entra As String = "C:\TS\FacturaUpdate\" + objFile.Name.ToString
                    Dim integ As String = "C:\TS\FacturaUpdate\Integration\Invoice.xml"
                    CancelaFactura()
                    'File.Delete("C:\TS\FacturaUpdate\Integration\Invoice.xml")
                    'File.Move(integ, temp)
                End If

            Next
        Catch ex As Exception

        End Try
    End Sub
    Public Sub OrdenesToma()
        Dim objFSO As Object = CreateObject("Scripting.FileSystemObject")
        Dim objSubFolder As Object = "C:\TS\Orden\"
        Dim objFolder As Object = objFSO.GetFolder(objSubFolder)
        Dim colFiles As Object = objFolder.Files

        For Each objFile In colFiles

            If existeOrden() = 0 Then
                Dim entra As String = "C:\TS\Orden\" + objFile.Name.ToString
                Dim integ As String = "C:\TS\Orden\Integration\Order.xml"
                Dim temp As String = "C:\TS\Orden\temp\Order" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml"
                File.Copy(entra, integ)
                File.Move(entra, temp)
            ElseIf existeOrden() = 1 Then
                Exit Sub
            End If
        Next
    End Sub
    Public Sub TrasladosToma()
        Dim objFSO As Object = CreateObject("Scripting.FileSystemObject")
        Dim objSubFolder As Object = "C:\TS\Traslados\"
        Dim objFolder As Object = objFSO.GetFolder(objSubFolder)
        Dim colFiles As Object = objFolder.Files

        For Each objFile In colFiles

            If existeTraslado() = 0 Then
                Dim entra As String = "C:\TS\Traslados\" + objFile.Name.ToString
                Dim integ As String = "C:\TS\Traslados\Integration\Traslado.xml"
                Dim temp As String = "C:\TS\Traslados\temp\Traslado" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml"
                File.Copy(entra, integ)
                File.Move(entra, temp)
            ElseIf existeTraslado() = 1 Then
                Exit Sub
            End If
        Next
    End Sub
    Public Sub NcreditoToma()
        Dim objFSO As Object = CreateObject("Scripting.FileSystemObject")
        Dim objSubFolder As Object = "C:\TS\NotaCredito\"
        Dim objFolder As Object = objFSO.GetFolder(objSubFolder)
        Dim colFiles As Object = objFolder.Files

        For Each objFile In colFiles

            If existeNcredito() = 0 Then
                Dim entra As String = "C:\TS\NotaCredito\" + objFile.Name.ToString
                Dim integ As String = "C:\TS\NotaCredito\Integration\NCredito.xml"
                Dim temp As String = "C:\TS\NotaCredito\temp\NCredito" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml"
                File.Copy(entra, integ)
                File.Move(entra, temp)
            ElseIf existeTraslado() = 1 Then
                Exit Sub
            End If
        Next
    End Sub
    Public Sub PacienteToma()
        Dim objFSO As Object = CreateObject("Scripting.FileSystemObject")
        Dim objSubFolder As Object = "C:\TS\Paciente\"
        Dim objFolder As Object = objFSO.GetFolder(objSubFolder)
        Dim colFiles As Object = objFolder.Files

        For Each objFile In colFiles

            If existePaciente() = 0 Then
                Dim entra As String = "C:\TS\Paciente\" + objFile.Name.ToString
                Dim integ As String = "C:\TS\Paciente\Integration\Paciente.xml"
                Dim temp As String = "C:\TS\Paciente\temp\Paciente" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml"
                File.Copy(entra, integ)
                File.Move(entra, temp)
            ElseIf existeTraslado() = 1 Then
                Exit Sub
            End If
        Next
    End Sub
    Public Sub PChequeToma()
        Dim objFSO As Object = CreateObject("Scripting.FileSystemObject")
        Dim objSubFolder As Object = "C:\TS\Pagos_Cheque\"
        Dim objFolder As Object = objFSO.GetFolder(objSubFolder)
        Dim colFiles As Object = objFolder.Files

        For Each objFile In colFiles

            If existePCheque() = 0 Then
                Dim entra As String = "C:\TS\Pagos_Cheque\" + objFile.Name.ToString
                Dim integ As String = "C:\TS\Pagos_Cheque\Integration\PaymentCHQ.xml"
                Dim temp As String = "C:\TS\Pagos_Cheque\temp\PaymentCHQ" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml"
                File.Copy(entra, integ)
                File.Move(entra, temp)
            ElseIf existeTraslado() = 1 Then
                Exit Sub
            End If
        Next
    End Sub
    Public Sub PCreditoToma()
        Dim objFSO As Object = CreateObject("Scripting.FileSystemObject")
        Dim objSubFolder As Object = "C:\TS\Pagos_Credito\"
        Dim objFolder As Object = objFSO.GetFolder(objSubFolder)
        Dim colFiles As Object = objFolder.Files

        For Each objFile In colFiles

            If existePCredito() = 0 Then
                Dim entra As String = "C:\TS\Pagos_Credito\" + objFile.Name.ToString
                Dim integ As String = "C:\TS\Pagos_Credito\Integration\PaymentCC.xml"
                Dim temp As String = "C:\TS\Pagos_Credito\temp\PaymentCC" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml"
                File.Copy(entra, integ)
                File.Move(entra, temp)
            ElseIf existeTraslado() = 1 Then
                Exit Sub
            End If
        Next
    End Sub
    Public Sub PEfectivoToma()
        Dim objFSO As Object = CreateObject("Scripting.FileSystemObject")
        Dim objSubFolder As Object = "C:\TS\Pagos_Efectivo\"
        Dim objFolder As Object = objFSO.GetFolder(objSubFolder)
        Dim colFiles As Object = objFolder.Files

        For Each objFile In colFiles

            If existePEfectivo() = 0 Then
                Dim entra As String = "C:\TS\Pagos_Efectivo\" + objFile.Name.ToString
                Dim integ As String = "C:\TS\Pagos_Efectivo\Integration\PaymentEF.xml"
                Dim temp As String = "C:\TS\Pagos_Efectivo\temp\PaymentEF" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml"
                File.Copy(entra, integ)
                File.Move(entra, temp)
            ElseIf existeTraslado() = 1 Then
                Exit Sub
            End If
        Next
    End Sub
    Public Sub PTransferToma()
        Dim objFSO As Object = CreateObject("Scripting.FileSystemObject")
        Dim objSubFolder As Object = "C:\TS\Pagos_Transfer\"
        Dim objFolder As Object = objFSO.GetFolder(objSubFolder)
        Dim colFiles As Object = objFolder.Files

        For Each objFile In colFiles

            If existePTransfer() = 0 Then
                Dim entra As String = "C:\TS\Pagos_Transfer\" + objFile.Name.ToString
                Dim integ As String = "C:\TS\Pagos_Transfer\Integration\PaymentTSFR.xml"
                Dim temp As String = "C:\TS\Pagos_Transfer\temp\PaymentTSFR" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml"
                File.Copy(entra, integ)
                File.Move(entra, temp)
            ElseIf existeTraslado() = 1 Then
                Exit Sub
            End If
        Next
    End Sub
    Public Sub PagosToma()
        Try


            Dim objFSO As Object = CreateObject("Scripting.FileSystemObject")
            Dim objSubFolder As Object = "C:\TS\Pagos\"
            Dim objFolder As Object = objFSO.GetFolder(objSubFolder)
            Dim colFiles As Object = objFolder.Files
            For Each objFile In colFiles

                If existePagos() = 0 Then
                    Dim entra As String = "C:\TS\Pagos\" + objFile.Name.ToString
                    Dim integ As String = "C:\TS\Pagos\Integration\Payment.xml"
                    File.Move(entra, integ)
                    PagoChequeIntegration()
                    'File.Delete("C:\TS\FacturaUpdate\Integration\Invoice.xml")
                    'File.Move(integ, temp)
                ElseIf existePagos() = 1 Then
                    Dim entra As String = "C:\TS\Pagos\" + objFile.Name.ToString
                    Dim integ As String = "C:\TS\Pagos\Integration\Payment.xml"
                    PagoChequeIntegration()
                    'File.Delete("C:\TS\FacturaUpdate\Integration\Invoice.xml")
                    'File.Move(integ, temp)
                End If

            Next
        Catch ex As Exception

        End Try
    End Sub



    Private Function existeItem()
        If My.Computer.FileSystem.FileExists("C:\TS\Items\Integration\Item.xml") Then
            'ListBox1.Items.Add("Si Existe")
            Return 1
        Else
            'ListBox1.Items.Add("No Existe")
            Return 0
        End If
    End Function
    Private Function existeFactura()
        If My.Computer.FileSystem.FileExists("C:\TS\Factura\Integration\Invoice.xml") Then
            'ListBox1.Items.Add("Si Existe")
            Return 1
        Else
            'ListBox1.Items.Add("No Existe")
            Return 0
        End If
    End Function
    Private Function existecFactura()
        If My.Computer.FileSystem.FileExists("C:\TS\FacturaUpdate\Integration\Invoice.xml") Then
            'ListBox1.Items.Add("Si Existe")
            Return 1
        Else
            'ListBox1.Items.Add("No Existe")
            Return 0
        End If
    End Function
    Private Function existeCliente()
        If My.Computer.FileSystem.FileExists("C:\TS\Cliente\Integration\Cliente.xml") Then
            'ListBox1.Items.Add("Si Existe")
            Return 1
        Else
            'ListBox1.Items.Add("No Existe")
            Return 0
        End If
    End Function
    Private Function existeOrden()
        If My.Computer.FileSystem.FileExists("C:\TS\Orden\Integration\Order.xml") Then
            'ListBox1.Items.Add("Si Existe")
            Return 1
        Else
            'ListBox1.Items.Add("No Existe")
            Return 0
        End If
    End Function
    Private Function existeTraslado()
        If My.Computer.FileSystem.FileExists("C:\TS\Traslados\Integration\Traslado.xml") Then
            'ListBox1.Items.Add("Si Existe")
            Return 1
        Else
            'ListBox1.Items.Add("No Existe")
            Return 0
        End If
    End Function
    Private Function existeNcredito()
        If My.Computer.FileSystem.FileExists("C:\TS\NotaCredito\Integration\NCredito.xml") Then
            'ListBox1.Items.Add("Si Existe")
            Return 1
        Else
            'ListBox1.Items.Add("No Existe")
            Return 0
        End If
    End Function
    Private Function existePaciente()
        If My.Computer.FileSystem.FileExists("C:\TS\Paciente\Integration\Paciente.xml") Then
            'ListBox1.Items.Add("Si Existe")
            Return 1
        Else
            'ListBox1.Items.Add("No Existe")
            Return 0
        End If
    End Function
    Private Function existePCheque()
        If My.Computer.FileSystem.FileExists("C:\TS\Pagos_Cheque\Integration\PaymentCHQ.xml") Then
            'ListBox1.Items.Add("Si Existe")
            Return 1
        Else
            'ListBox1.Items.Add("No Existe")
            Return 0
        End If
    End Function
    Private Function existePCredito()
        If My.Computer.FileSystem.FileExists("C:\TS\Pagos_Credito\Integration\PaymentCC.xml") Then
            'ListBox1.Items.Add("Si Existe")
            Return 1
        Else
            'ListBox1.Items.Add("No Existe")
            Return 0
        End If
    End Function
    Private Function existePEfectivo()
        If My.Computer.FileSystem.FileExists("C:\TS\Pagos_Efectivo\Integration\PaymentEF.xml") Then
            'ListBox1.Items.Add("Si Existe")
            Return 1
        Else
            'ListBox1.Items.Add("No Existe")
            Return 0
        End If
    End Function
    Private Function existePTransfer()
        If My.Computer.FileSystem.FileExists("C:\TS\Pagos_Transfer\Integration\PaymentTSFR.xml") Then
            'ListBox1.Items.Add("Si Existe")
            Return 1
        Else
            'ListBox1.Items.Add("No Existe")
            Return 0
        End If
    End Function
    Private Function existePagos()
        If My.Computer.FileSystem.FileExists("C:\TS\Pagos\Integration\Payment.xml") Then
            'ListBox1.Items.Add("Si Existe")
            Return 1
        Else
            'ListBox1.Items.Add("No Existe")
            Return 0
        End If
    End Function


    Private Sub CancelaFactura()

        Dim oReturn As Integer = -1
        Dim oError As Integer = 0
        Dim errMsg As String = ""
        Try
            Dim entra As String = "C:\TS\FacturaUpdate\Integration\Invoice.xml"
            Dim sale As String = "C:\TS\FacturaUpdate\temp\Invoice" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml"
            Dim Xml As XmlDocument = New XmlDocument()
            Xml.Load(entra)
            Dim ArticleList As XmlNodeList = Xml.SelectNodes("invoice/document")
            For Each xnDoc As XmlNode In ArticleList
                If MakeConnectionSAP() Then
                    Dim Invoice As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                    Dim oInvoice2 As SAPbobsCOM.Documents
                    Invoice.GetByKey(xnDoc.SelectSingleNode("docentry").InnerText)
                    oInvoice2 = Invoice.CreateCancellationDocument()
                    If oInvoice2.Add() = 0 Then
                        File.Move(entra, sale)
                    End If
                End If
            Next
        Catch ex As Exception
            Dim entra As String = "C:\TS\FacturaUpdate\Integration\Invoice.xml"
            Dim sale As String = "C:\TS\FacturaUpdate\temp\ErrorInvoice" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml"
            File.Move(entra, sale)
        End Try
    End Sub


    Private Sub PagoChequeIntegration()
        Dim oReturn As Integer = -1
        Dim oError As Integer = 0
        Dim errMsg As String = ""
        Dim sql As String
        Try
            Dim entra As String = "C:\TS\Pagos\Integration\Payment.xml"
            Dim sale As String = "C:\TS\Pagos\temp\Payment" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml"
            Dim Xml As XmlDocument = New XmlDocument()
            Xml.Load(entra)
            Dim ArticleList As XmlNodeList = Xml.SelectNodes("body/document")
            For Each xnDoc As XmlNode In ArticleList
                If MakeConnectionSAP() Then
                    Dim oPayment As SAPbobsCOM.Payments = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)

                    oPayment.DocNum = xnDoc.SelectSingleNode("tipodepago").InnerText
                    oPayment.DocNum = xnDoc.SelectSingleNode("pagoacuenta").InnerText
                    oPayment.DocNum = xnDoc.SelectSingleNode("docnum").InnerText
                    oPayment.Series = xnDoc.SelectSingleNode("series").InnerText
                    oPayment.DocDate = xnDoc.SelectSingleNode("docdate").InnerText
                    oPayment.CardCode = xnDoc.SelectSingleNode("cardcode").InnerText
                    oPayment.DocCurrency = xnDoc.SelectSingleNode("doccurrency").InnerText
                    oPayment.DocType = oPayment.DocType.rCustomer
                    'oPayment.DocType = xnDoc.SelectSingleNode("remarks").InnerText

                    Dim CatNodesLines As XmlNodeList = xnDoc.SelectNodes("document_lines/line")
                    For Each xnDetLines As XmlNode In CatNodesLines

                        sql = ("select top 1 DocEntry from oinv where DocNum = " + xnDetLines.SelectSingleNode("docentry").InnerText)
                        Dim oRecordSet As SAPbobsCOM.Recordset
                        oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery(sql)
                        If oRecordSet.RecordCount > 0 Then
                            oPayment.Invoices.DocEntry = oRecordSet.Fields.Item(0).Value
                        End If
                        oPayment.Invoices.InvoiceType = "13"
                        oPayment.Invoices.SumApplied = xnDetLines.SelectSingleNode("sumapplied").InnerText
                        oPayment.Invoices.Add()
                    Next
                    Dim CatNodesPay As XmlNodeList = xnDoc.SelectNodes("document_lines/pay")
                    For Each xnDetPay As XmlNode In CatNodesPay
                        oPayment.Checks.CheckNumber = xnDetPay.SelectSingleNode("checknumber").InnerText
                        oPayment.Checks.BankCode = xnDetPay.SelectSingleNode("bankcode").InnerText
                        oPayment.Checks.AccounttNum = xnDetPay.SelectSingleNode("accounttnum").InnerText
                        oPayment.Checks.CheckSum = xnDetPay.SelectSingleNode("checksum").InnerText
                        oPayment.Checks.CheckAccount = xnDetPay.SelectSingleNode("checkaccount").InnerText
                        oPayment.Checks.Add()
                    Next


                    oPayment.Remarks = xnDoc.SelectSingleNode("remarks").InnerText
                    oReturn = oPayment.Add()
                    If oReturn <> 0 Then
                        oCompany.GetLastError(oError, errMsg)
                        MsgBox(errMsg)
                    End If
                End If
            Next
            File.Move(entra, sale)
        Catch ex As Exception
            Dim entra As String = "C:\TS\Pagos\Integration\Payment.xml"
            Dim sale As String = "C:\TS\Pagos\temp\Payment" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml"
            File.Move(entra, sale)
        End Try
    End Sub
End Class
