Imports System.Windows.Forms
Imports System.IO
Imports System.Xml
Imports System.Threading
Public Class Form1
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

            If existeTraslado() = 0 Then
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

            If existeTraslado() = 0 Then
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

            If existeTraslado() = 0 Then
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

            If existeTraslado() = 0 Then
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

            If existeTraslado() = 0 Then
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

            If existeTraslado() = 0 Then
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
End Class
