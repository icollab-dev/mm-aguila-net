Imports System.Linq
Imports System.Threading
Imports System.Web
Imports System.Web.Services
'Imports BLL
'Imports BLL.MModel
'Imports DELL.MModel

<WebService([Namespace]:="http://localhost/")>
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<System.ComponentModel.ToolboxItem(False)>
Public Class WebServiceGeneral
    'Inherits MModel

    <WebMethod>
    Public Function ESTSingle(ByVal input As ESTInput) As ESTOutput
        Dim estOutput As ESTOutput = New ESTOutput()
        Dim exeProcess As Process = New Process()
        Dim startInfo As ProcessStartInfo = New ProcessStartInfo()
        Dim cccmModelEstPtr As CCCMModel_EST_Ptr = New CCCMModel_EST_Ptr()
        startInfo.FileName = ConfigurationManager.AppSettings("ESTFileName")
        startInfo.WorkingDirectory = HttpContext.Current.Server.MapPath(ConfigurationManager.AppSettings("ESTRelativePath"))
        startInfo.CreateNoWindow = False
        startInfo.WindowStyle = ProcessWindowStyle.Hidden
        Dim cccmModelEst As ESTOutput

        Dim _bll As New BLL.MModel

        Try
            exeProcess = Process.Start(startInfo)
            exeProcess.WaitForInputIdle()
            Thread.Sleep(10)
            Dim ptr As CCCMModel_EST_Ptr = Me.SetPtr_EST()
            cccmModelEst = Me.GetCCCMModel_EST(exeProcess, input, ptr)
            exeProcess.WaitForInputIdle()
            Thread.Sleep(10)
            _bll.Doclick(ptr.ExitButton)
            Thread.Sleep(10)

            If Not exeProcess.HasExited Then
                exeProcess.WaitForExit(10)
                exeProcess.Kill()
            End If

        Catch ex1 As Exception
            estOutput = New ESTOutput()

            Try

                If Not exeProcess.HasExited Then
                    exeProcess.WaitForExit(10)
                    exeProcess.Kill()
                End If

            Catch ex2 As Exception
                Throw New Exception(ex2.Message)
            End Try

            Throw New Exception(ex1.Message)
        End Try

        Return cccmModelEst
    End Function

    <WebMethod>
    Public Function ESTBatch(ByVal inputs As List(Of ESTInput)) As List(Of ESTOutput)
        Dim exeProcess As Process = New Process()
        Dim startInfo As ProcessStartInfo = New ProcessStartInfo()
        Dim estOutputList1 As List(Of ESTOutput) = New List(Of ESTOutput)()
        Dim cccmModelEstPtr As CCCMModel_EST_Ptr = New CCCMModel_EST_Ptr()
        startInfo.FileName = ConfigurationManager.AppSettings("ESTFileName")
        startInfo.WorkingDirectory = HttpContext.Current.Server.MapPath(ConfigurationManager.AppSettings("ESTRelativePath"))
        startInfo.CreateNoWindow = False
        startInfo.WindowStyle = ProcessWindowStyle.Hidden

        Dim _bll As New BLL.MModel

        Try
            exeProcess = Process.Start(startInfo)
            exeProcess.WaitForInputIdle()
            Thread.Sleep(10)
            Dim ptr As CCCMModel_EST_Ptr = Me.SetPtr_EST()

            For Each input As ESTInput In inputs
                estOutputList1.Add(Me.GetCCCMModel_EST(exeProcess, input, ptr))
            Next

            exeProcess.WaitForInputIdle()
            Thread.Sleep(10)
            _bll.Doclick(ptr.ExitButton)
            exeProcess.WaitForInputIdle()
            Thread.Sleep(10)

            If Not exeProcess.HasExited Then
                exeProcess.WaitForExit(10)
                exeProcess.Kill()
            End If

        Catch ex1 As Exception
            Dim estOutputList2 As List(Of ESTOutput) = New List(Of ESTOutput)()

            Try

                If Not exeProcess.HasExited Then
                    exeProcess.WaitForExit(10)
                    exeProcess.Kill()
                End If

            Catch ex2 As Exception
                Throw New Exception(ex2.Message)
            End Try

            Throw New Exception(ex1.Message)
        End Try

        Return estOutputList1
    End Function

    Private Function SetPtr_EST() As CCCMModel_EST_Ptr
        Dim cccmModelEstPtr As CCCMModel_EST_Ptr = New CCCMModel_EST_Ptr()

        cccmModelEstPtr.MainWindow = MModel_BLL.FindWindow(CStr(Nothing), "Modelo Matematico Ver5.0.1")
        cccmModelEstPtr.ExecuteButton = MModel_BLL.FindWindowEx(cccmModelEstPtr.MainWindow, New IntPtr(), "ThunderRT6CommandButton", "Calculo")
        cccmModelEstPtr.ExitButton = MModel_BLL.FindWindowEx(cccmModelEstPtr.MainWindow, New IntPtr(), "ThunderRT6CommandButton", "Salir")
        cccmModelEstPtr.FrameInputData = MModel_BLL.FindWindowEx(cccmModelEstPtr.MainWindow, New IntPtr(), "ThunderRT6Frame", "Datos de la Entrada")

        cccmModelEstPtr.CurrentCapacity = MModel_BLL.GetDlgItem(cccmModelEstPtr.FrameInputData, 37)
        cccmModelEstPtr.DryBulbTemp = MModel_BLL.GetDlgItem(cccmModelEstPtr.FrameInputData, 35)
        cccmModelEstPtr.AtmosphericPressure = MModel_BLL.GetDlgItem(cccmModelEstPtr.FrameInputData, 34)
        cccmModelEstPtr.HumidityRelative = MModel_BLL.GetDlgItem(cccmModelEstPtr.FrameInputData, 33)
        cccmModelEstPtr.CalorificPower = MModel_BLL.GetDlgItem(cccmModelEstPtr.FrameInputData, 32)
        cccmModelEstPtr.SeaWaterTemp = MModel_BLL.GetDlgItem(cccmModelEstPtr.FrameInputData, 31)
        cccmModelEstPtr.PowerFactor = MModel_BLL.GetDlgItem(cccmModelEstPtr.FrameInputData, 30)
        cccmModelEstPtr.FrameCalculationResult = MModel_BLL.FindWindowEx(cccmModelEstPtr.MainWindow, New IntPtr(), "ThunderRT6Frame", "Resultad del Calculo")
        cccmModelEstPtr.CurrentPlantLoad = MModel_BLL.GetDlgItem(cccmModelEstPtr.FrameCalculationResult, 28)
        cccmModelEstPtr.NetCapacity = MModel_BLL.GetDlgItem(cccmModelEstPtr.FrameCalculationResult, 27)
        cccmModelEstPtr.SummerCapacity = MModel_BLL.GetDlgItem(cccmModelEstPtr.FrameCalculationResult, 26)
        cccmModelEstPtr.CTUNGLoad = MModel_BLL.GetDlgItem(cccmModelEstPtr.FrameCalculationResult, 25)
        cccmModelEstPtr.CTUNG = MModel_BLL.GetDlgItem(cccmModelEstPtr.FrameCalculationResult, 24)
        cccmModelEstPtr.CTOV = MModel_BLL.GetDlgItem(cccmModelEstPtr.FrameCalculationResult, 23)
        cccmModelEstPtr.FramePowerCorrection = MModel_BLL.FindWindowEx(cccmModelEstPtr.MainWindow, New IntPtr(), "ThunderRT6Frame", "Correccion de Potencia")
        cccmModelEstPtr.Cpw_AT = MModel_BLL.GetDlgItem(cccmModelEstPtr.FramePowerCorrection, 21)
        cccmModelEstPtr.Cpw_BP = MModel_BLL.GetDlgItem(cccmModelEstPtr.FramePowerCorrection, 20)
        cccmModelEstPtr.Cpw_RH = MModel_BLL.GetDlgItem(cccmModelEstPtr.FramePowerCorrection, 19)
        cccmModelEstPtr.Cpw_PCI = MModel_BLL.GetDlgItem(cccmModelEstPtr.FramePowerCorrection, 18)
        cccmModelEstPtr.Cpw_CW = MModel_BLL.GetDlgItem(cccmModelEstPtr.FramePowerCorrection, 17)
        cccmModelEstPtr.Cpw_FP = MModel_BLL.GetDlgItem(cccmModelEstPtr.FramePowerCorrection, 16)
        cccmModelEstPtr.FrameCTUNGCorrection = MModel_BLL.FindWindowEx(cccmModelEstPtr.MainWindow, New IntPtr(), "ThunderRT6Frame", "Correccion de CTUNG y CTOV")
        cccmModelEstPtr.Chr_AT = MModel_BLL.GetDlgItem(cccmModelEstPtr.FrameCTUNGCorrection, 7)
        cccmModelEstPtr.Chr_BP = MModel_BLL.GetDlgItem(cccmModelEstPtr.FrameCTUNGCorrection, 6)
        cccmModelEstPtr.Chr_RH = MModel_BLL.GetDlgItem(cccmModelEstPtr.FrameCTUNGCorrection, 5)
        cccmModelEstPtr.Chr_PCI = MModel_BLL.GetDlgItem(cccmModelEstPtr.FrameCTUNGCorrection, 4)
        cccmModelEstPtr.Chr_CW = MModel_BLL.GetDlgItem(cccmModelEstPtr.FrameCTUNGCorrection, 3)
        cccmModelEstPtr.Chr_FP = MModel_BLL.GetDlgItem(cccmModelEstPtr.FrameCTUNGCorrection, 2)
        Return cccmModelEstPtr
    End Function

    Private Function GetCCCMModel_EST(ByVal exeProcess As Process, ByVal input As ESTInput, ByVal ptr As CCCMModel_EST_Ptr) As ESTOutput
        Dim estOutput As ESTOutput = New ESTOutput()

        Dim _dell As New DELL.MModel

        MModel_BLL.SetWindowTextRaw(ptr.CurrentCapacity, input.currentCapacity.ToString())
        MModel_BLL.SetWindowTextRaw(ptr.DryBulbTemp, input.dryBulbTemp.ToString())
        MModel_BLL.SetWindowTextRaw(ptr.AtmosphericPressure, input.atmosphericPressure.ToString())
        MModel_BLL.SetWindowTextRaw(ptr.HumidityRelative, input.humidityRelative.ToString())
        MModel_BLL.SetWindowTextRaw(ptr.CalorificPower, input.calorificPower.ToString())
        MModel_BLL.SetWindowTextRaw(ptr.SeaWaterTemp, input.seaWaterTemp.ToString())
        MModel_BLL.SetWindowTextRaw(ptr.PowerFactor, input.powerFactor.ToString())
        MModel_BLL.Doclick(ptr.ExecuteButton)
        Thread.Sleep(10)
        MModel_BLL.Doclick(ptr.ExecuteButton)
        exeProcess.WaitForInputIdle()
        Thread.Sleep(10)
        Dim windowTextRaw1 As String = MModel_BLL.GetWindowTextRaw(ptr.CurrentPlantLoad)
        Dim windowTextRaw2 As String = MModel_BLL.GetWindowTextRaw(ptr.NetCapacity)
        Dim windowTextRaw3 As String = MModel_BLL.GetWindowTextRaw(ptr.SummerCapacity)
        Dim windowTextRaw4 As String = MModel_BLL.GetWindowTextRaw(ptr.CTUNGLoad)
        Dim windowTextRaw5 As String = MModel_BLL.GetWindowTextRaw(ptr.CTUNG)
        Dim windowTextRaw6 As String = MModel_BLL.GetWindowTextRaw(ptr.CTOV)
        Dim windowTextRaw7 As String = MModel_BLL.GetWindowTextRaw(ptr.Cpw_AT)
        Dim windowTextRaw8 As String = MModel_BLL.GetWindowTextRaw(ptr.Cpw_BP)
        Dim windowTextRaw9 As String = MModel_BLL.GetWindowTextRaw(ptr.Cpw_RH)
        Dim windowTextRaw10 As String = MModel_BLL.GetWindowTextRaw(ptr.Cpw_PCI)
        Dim windowTextRaw11 As String = MModel_BLL.GetWindowTextRaw(ptr.Cpw_CW)
        Dim windowTextRaw12 As String = MModel_BLL.GetWindowTextRaw(ptr.Cpw_FP)
        Dim windowTextRaw13 As String = MModel_BLL.GetWindowTextRaw(ptr.Chr_AT)
        Dim windowTextRaw14 As String = MModel_BLL.GetWindowTextRaw(ptr.Chr_BP)
        Dim windowTextRaw15 As String = MModel_BLL.GetWindowTextRaw(ptr.Chr_RH)
        Dim windowTextRaw16 As String = MModel_BLL.GetWindowTextRaw(ptr.Chr_PCI)
        Dim windowTextRaw17 As String = MModel_BLL.GetWindowTextRaw(ptr.Chr_CW)
        Dim windowTextRaw18 As String = MModel_BLL.GetWindowTextRaw(ptr.Chr_FP)
        estOutput.idCincominutal = input.idCincominutal
        estOutput.currentCapacity = input.currentCapacity
        estOutput.dryBulbTemp = input.dryBulbTemp
        estOutput.atmosphericPressure = input.atmosphericPressure
        estOutput.humidityRelative = input.humidityRelative
        estOutput.calorificPower = input.calorificPower
        estOutput.seaWaterTemp = input.seaWaterTemp
        estOutput.powerFactor = input.powerFactor
        estOutput.CurrentPlantLoad = Decimal.Parse(windowTextRaw1)
        estOutput.NetCapacity = Decimal.Parse(windowTextRaw2)
        estOutput.SummerCapacity = Decimal.Parse(windowTextRaw3)
        estOutput.CTUNGLoad = windowTextRaw4
        estOutput.CTUNG = windowTextRaw5
        estOutput.CTOV = windowTextRaw6
        estOutput.Cpw_AT = Decimal.Parse(windowTextRaw7)
        estOutput.Cpw_BP = Decimal.Parse(windowTextRaw8)
        estOutput.Cpw_RH = Decimal.Parse(windowTextRaw9)
        estOutput.Cpw_PCI = Decimal.Parse(windowTextRaw10)
        estOutput.Cpw_CW = Decimal.Parse(windowTextRaw11)
        estOutput.Cpw_FP = Decimal.Parse(windowTextRaw12)
        estOutput.Chr_AT = Decimal.Parse(windowTextRaw13)
        estOutput.Chr_BP = Decimal.Parse(windowTextRaw14)
        estOutput.Chr_RH = Decimal.Parse(windowTextRaw15)
        estOutput.Chr_PCI = Decimal.Parse(windowTextRaw16)
        estOutput.Chr_CW = Decimal.Parse(windowTextRaw17)
        estOutput.Chr_FP = Decimal.Parse(windowTextRaw18)
        Return estOutput
    End Function
End Class

