Imports Microsoft.VisualBasic
Imports Forms = System.Windows.Forms
Imports IO = System.IO
Imports System
Imports System.Runtime.InteropServices
Imports Tekla.BC.Plugin
Imports Tekla.Structures
Imports Tekla.Structures.Model

Namespace Tekla.BC.Structures.Macros

    Class ExcelMacro

        Const DEBUG As Boolean = True
        Const SHOW_EXCEL As Boolean = True
        Const RUN_MACRO As Boolean = True
        Const STORE_RESULTS As Boolean = True

        Shared Iface As CDelegate = Nothing
        
#Region "Original Basic Get and Set Functions"
        Private Shared Function GetDoubleValue _
            (src As String, ByRef value As Double) As Integer
            Return Iface.GetDoubleValue(src, value)
        End Function

        Private Shared Function SetDoubleValue _
            (src As String, value As Double) As Integer
            Return Iface.SetDoubleValue(src, value)
        End Function


        Private Shared Function GetIntValue _
            (src As String, ByRef value As Integer) As Integer
            Return Iface.GetIntValue(src, value)
        End Function


        Private Shared Function SetIntValue _
            (src As String, value As Integer) As Integer
            Return Iface.SetIntValue(src, value)
        End Function


        Private Shared Function GetStringValue _
            (src As String, ByRef value As String) As Integer
            Return Iface.GetStringValue(src, value)
        End Function


        Private Shared Function SetStringValue _
            (src As String, value As String) As Integer
            Return Iface.SetStringValue(src, value)
        End Function


        Private Shared Function GetJointStructure _
            (ByRef Joint As dotJoint_t) As Integer
            Return Iface.GetJointStructure(Joint)
        End Function

        Private Shared Function GetComponentName _
            (Id As Integer, ByRef value As String) As Integer
            Return Iface.GetComponentName(Id, value)
        End Function

        Private Shared Function GetCurrentModelDirectory _
            (ByRef value As String) As Integer
            Return Iface.GetCurrentModelDirectory(value)
        End Function

        Private Shared Function GetWorkBookDirectory _
            (workBookName As String, storageDir As String) As String
            Dim value As String
            Dim res = Iface.GetWorkBookDirectory(workBookName, storageDir, value)
            If res = 0 Then
                value = ""
            End If
            Return value
        End Function

        Private Shared Function GetPartCoordinateSystem _
            (Id As Integer, ByRef X As dotPoint_t, ByRef Y As dotPoint_t, _
            ByRef Z As dotPoint_t) As Integer
            Return Iface.GetPartCoordinateSystem(Id, X, Y, Z)
        End Function


        Private Shared Function GetPolybeamCoordinateSystem _
            (Id As Integer, SubId As Integer, Chamfered As Integer, _
            ByRef X As dotPoint_t, ByRef Y As dotPoint_t, ByRef Z As dotPoint_t) As Integer
            Return Iface.GetPolybeamCoordinateSystem(Id, SubId, Chamfered, X, Y, Z)
        End Function


        Private Shared Function GetPartExtrema _
            (PartId As Integer, ByRef Min As dotPoint_t, ByRef Max As dotPoint_t) As Integer
            Return Iface.GetPartExtrema(PartId, Min, Max)
        End Function


        Private Shared Function GetMaterialProperties _
            (MaterialName As String, ByRef Elasticity As Double, ByRef Poisson As Double) _
            As Integer
            Return Iface.GetMaterialProperties(MaterialName, Elasticity, Poisson)
        End Function


        Private Shared Function GetJointPrimaryFramingCondition _
            (ByRef Primary As dotFramingConditionOfPrimary_t) As Integer
            Return Iface.GetJointPrimaryFramingCondition(Primary)
        End Function


        Private Shared Function GetJointSecondaryFramingCondition _
            (Num As Integer, ByRef Secondary As dotFramingConditionOfSecondary_t) As Integer
            Return Iface.GetJointSecondaryFramingCondition(Num, Secondary)
        End Function


        Private Shared Function GetProfileData(Id As Integer, ByRef ProfileData As dotProfileData_t)
            Return Iface.GetProfileData(Id, ProfileData)
        End Function



        Private Shared Sub NAR(ByVal o As Object)
            Try
                System.Runtime.InteropServices.Marshal.ReleaseComObject(o)
            Catch
                'Forms.MessageBox.Show("NAR EXCEPTION", "NAR EXCEPTION")
            Finally
                o = Nothing
            End Try
        End Sub


        Private Shared Function getJointSecondariesCount() As Integer
            Dim Joint As dotJoint_t
            GetJointStructure(Joint)
            Return Joint.nSecondaries
        End Function

        Private Shared Function getJointNumber() As Integer
            Dim Joint As dotJoint_t
            GetJointStructure(Joint)
            Return Joint.JointNumber
        End Function

        Private Shared Function getComponentName() As String
            Dim Joint As dotJoint_t
            Dim Name As String
            GetJointStructure(Joint)
            getComponentName(Joint.Id, Name)
            Return Name
        End Function
        Private Shared Function getJointStructVariable(Name As String) As Object
            Dim Joint As dotJoint_t

            Name = Name.ToLower()
            GetJointStructure(Joint)

            If Name.Equals("id") Then
				Dim jGuid As String
				Iface.GetGuid(Convert.toInt32(Joint.Id), jGuid)
                Return jGuid
            ElseIf Name.Equals("group") Then
                Return Joint.Group
            ElseIf Name.Equals("flags") Then
                Return Joint.Flags
            ElseIf Name.Equals("jointnumber") Then
                Return Joint.JointNumber
            ElseIf Name.Equals("up.x") Then
                Return Joint.Up.x
            ElseIf Name.Equals("up.y") Then
                Return Joint.Up.y
            ElseIf Name.Equals("up.z") Then
                Return Joint.Up.z
            ElseIf Name.Equals("modeldirectory") Then
                Dim modelDir As String
                GetCurrentModelDirectory(modelDir)
                Return modelDir
            End If
            Return Nothing
        End Function
        Private Shared Function getPartCoordinateSystem(JointId As Integer, Name As String, ByRef Value As Double) As Boolean
            Dim pX As dotPoint_t
            Dim pY As dotPoint_t
            Dim pZ As dotPoint_t
            Dim sel As dotPoint_t
            Dim names As String()

            getPartCoordinateSystem(JointId, pX, pY, pZ)

            names = Name.Split(".".ToCharArray())
            If names.Length <> 2 Then
                Return False
            End If

            names(0) = names(0).ToLower()
            names(1) = names(1).ToLower()

            If names(0).Equals("x") Then
                sel = pX
            ElseIf names(0).Equals("y") Then
                sel = pY
            ElseIf names(0).Equals("z") Then
                sel = pZ
            Else
                Return False
            End If

            If names(1).Equals("x") Then
                Value = sel.x
                Return True
            ElseIf names(1).Equals("y") Then
                Value = sel.y
                Return True
            ElseIf names(1).Equals("z") Then
                Value = sel.z
                Return True
            End If
            Return False
        End Function

        Private Shared Function getPartExtrema(JointId As Integer, Name As String, ByRef Value As Double) As Boolean
            Dim pMin As dotPoint_t
            Dim pMax As dotPoint_t
            Dim sel As dotPoint_t
            Dim names As String()

            getPartExtrema(JointId, pMin, pMax)

            names = Name.Split(".".ToCharArray())
            If names.Length <> 2 Then
                Return False
            End If

            names(0) = names(0).ToLower()
            names(1) = names(1).ToLower()

            If names(0).Equals("min") Then
                sel = pMin
            ElseIf names(0).Equals("max") Then
                sel = pMax
            Else
                Return False
            End If

            If names(1).Equals("x") Then
                Value = sel.x
                Return True
            ElseIf names(1).Equals("y") Then
                Value = sel.y
                Return True
            ElseIf names(1).Equals("z") Then
                Value = sel.z
                Return True
            End If
            Return False
        End Function


        Private Shared Function setVariable(name As String, value As Object) As Boolean
            If TypeOf value Is Double Then
                SetDoubleValue(name, value)
                Return True
            End If
            If TypeOf value Is Integer Then
                SetIntValue(name, value)
                Return True
            End If
            If TypeOf value Is String Then
                SetStringValue(name, value)
                Return True
            End If
            Return False
        End Function
#End Region

#Region "New Get Report and Set UDA Functions"
        Private Shared Function GetStringReportValue(ByVal id As Integer, ByVal ReportField As String) As String
            Dim beam As Beam
            Dim stringValue As String
            beam = New Beam
            beam.Identifier.ID = id
            stringValue = ""
            beam.GetReportProperty(ReportField, stringValue)
            Return stringValue
        End Function
        Private Shared Function GetIntegerReportValue(ByVal id As Integer, ByVal ReportField As String) As Integer
            Dim beam As Beam
            Dim integerValue As Integer
            beam = New Beam
            beam.Identifier.ID = id
            integerValue = 0
            beam.GetReportProperty(ReportField, integerValue)
            Return integerValue
        End Function
        Private Shared Function GetDoubleReportValue(ByVal id As Integer, ByVal ReportField As String) As Double
            Dim beam As Beam
            Dim doubleValue As Double
            beam = New Beam
            beam.Identifier.ID = id
            doubleValue = 0
            beam.GetReportProperty(ReportField, doubleValue)
            Return doubleValue
        End Function
        Private Shared Function SetStringUdaValue(ByVal id As Integer, ByVal UdaField As String, ByVal UdaValue As String) As Object
            Dim beam As Beam
            beam = New Beam
            beam.Identifier.ID = id
            beam.SetUserProperty(UdaField, UdaValue)
            Return Nothing
        End Function
        Private Shared Function SetIntegerUdaValue(ByVal id As Integer, ByVal UdaField As String, ByVal UdaValue As Integer) As Object
            Dim beam As Beam
            beam = New Beam
            beam.Identifier.ID = id
            beam.SetUserProperty(UdaField, UdaValue)
            Return Nothing
        End Function
        Private Shared Function SetDoubleUdaValue(ByVal id As Integer, ByVal UdaField As String, ByVal UdaValue As Double) As Object
            Dim beam As Beam
            beam = New Beam
            beam.Identifier.ID = id
            beam.SetUserProperty(UdaField, UdaValue)
            Return Nothing
        End Function
        Private Shared Function getJointVariable(Idx As Integer, Name As String) As Object
            Dim Joint As dotJoint_t
            Dim Id As Integer
            Dim Value As Object

            GetJointStructure(Joint)

            If Idx > Joint.nSecondaries Then
                Return Nothing
            End If

            If Idx = 0 Then
                Id = Joint.Primary
            Else
                Id = Joint.Secondaries(Idx - 1)
            End If

            If Name.StartsWith("RPT.STR.") Then
                Name = Name.Remove(0, 8)
                Return GetStringReportValue(Id, Name)
            End If

            If Name.StartsWith("RPT.INT.") Then
                Name = Name.Remove(0, 8)
                Return GetIntegerReportValue(Id, Name)
            End If

            If Name.StartsWith("RPT.DBL.") Then
                Name = Name.Remove(0, 8)
                Return GetDoubleReportValue(Id, Name)
            End If

            Name = Name.ToLower()

            ' handle PartCoordinates
            If GetPartCoordinateSystem(Id, Name, Value) Then
                Return Value
            End If

            ' handle PartExtrema
            If GetPartExtrema(Id, Name, Value) Then
                Return Value
            End If

            ' handle FramingCondition
            If Idx = 0 Then
                Dim FC As dotFramingConditionOfPrimary_t
                GetJointPrimaryFramingCondition(FC)
                If Name.Equals("id") Then
					Dim pGuid As String
					Iface.GetGuid(Convert.toInt32(Joint.Primary), pGuid)
                    Return pGuid
                ElseIf Name.Equals("type") Then
                    Return FC.Type
                ElseIf Name.Equals("name") Then
                    Return FC.Name
                ElseIf Name.Equals("profiletype") Then
                    Return FC.ProfileType
                End If
            Else
                Dim FC As dotFramingConditionOfSecondary_t
                GetJointSecondaryFramingCondition(Idx - 1, FC)
                If Name.Equals("id") Then
					Dim sGuid As String
					Iface.GetGuid(Convert.toInt32(Joint.Secondaries(Idx - 1)), sGuid)
                    Return sGuid
                ElseIf Name.Equals("type") Then
                    Return FC.Type
                ElseIf Name.Equals("name") Then
                    Return FC.Name
                ElseIf Name.Equals("profiletype") Then
                    Return FC.ProfileType
                ElseIf Name.Equals("skewangle") Then
                    Return FC.SkewAngle
                ElseIf Name.Equals("slopeangle") Then
                    Return FC.SlopeAngle
                ElseIf Name.Equals("anglecant") Then
                    Return FC.AngleCant
                ElseIf Name.Equals("offset") Then
                    Return FC.Offset
                ElseIf Name.Equals("shearforce") Then
                    Return FC.ShearForce
                ElseIf Name.Equals("axialforce") Then
                    Return FC.AxialForce
                ElseIf Name.Equals("bendingmoment") Then
                    Return FC.BendingMoment
                ElseIf Name.Equals("useudl") Then
                    Return FC.UseUDL
                ElseIf Name.Equals("udlpercent") Then
                    Return FC.UDLPercent
                ElseIf Name.Equals("connectioncode") Then
                    Return FC.ConnectionCode
                ElseIf Name.Equals("nearestconnectionplane") Then
                    Return FC.NearestConnectionPlane
                End If
            End If

            ' handle ProfileData
            Dim ProfileData As dotProfileData_t
            Dim PartId As Integer
            If Idx = 0 Then
                Dim FC As dotFramingConditionOfPrimary_t
                GetJointPrimaryFramingCondition(FC)
                PartId = FC.Id
            Else
                Dim FC As dotFramingConditionOfSecondary_t
                GetJointSecondaryFramingCondition(Idx - 1, FC)
                PartId = FC.Id
            End If

            GetProfileData(PartId, ProfileData)
            If Name.Equals("type") Then
                Return ProfileData.Type
            ElseIf Name.Equals("r1") Then
                Return ProfileData.R1
            ElseIf Name.Equals("h") Then
                Return ProfileData.H
            ElseIf Name.Equals("b") Then
                Return ProfileData.B
            ElseIf Name.Equals("s") Then
                Return ProfileData.S
            ElseIf Name.Equals("t") Then
                Return ProfileData.T
            ElseIf Name.Equals("r") Then
                Return ProfileData.R
            ElseIf Name.Equals("h2") Then
                Return ProfileData.H2
            ElseIf Name.Equals("b2") Then
                Return ProfileData.B2
            ElseIf Name.Equals("s2") Then
                Return ProfileData.S2
            ElseIf Name.Equals("t2") Then
                Return ProfileData.T2
            ElseIf Name.Equals("r2") Then
                Return ProfileData.R2
            ElseIf Name.Equals("wx") Then
                Return ProfileData.Wx
            ElseIf Name.Equals("wy") Then
                Return ProfileData.Wy
            ElseIf Name.Equals("fs") Then
                Return ProfileData.Fs
            ElseIf Name.Equals("ex") Then
                Return ProfileData.Ex
            ElseIf Name.Equals("ey") Then
                Return ProfileData.Ey
            ElseIf Name.Equals("profilename") Then
                Return ProfileData.ProfileName
            ElseIf Name.Equals("materialname") Then
                Return ProfileData.MaterialName
            End If

            'Forms.MessageBox.Show("Nothing")

            ' No match
            Return Nothing
        End Function
#End Region

        Private Shared Function GetOutputDirectory(ByVal Iface As CDelegate) As String
            'New optional output directory 
            Dim storageDir = String.Empty
            Dim fetchResult As Integer
            fetchResult = Iface.GetEnvironmentString("XS_EXTERNAL_DESIGN_STORAGE_PATH", storageDir)
            If fetchResult = 0 Then
                GetCurrentModelDirectory(storageDir)
                storageDir = storageDir + "\exceldesign\"
            Else
                Dim model = New Model()
                Dim modelInfo = model.GetInfo
                storageDir = storageDir + modelInfo.ModelName.Substring(0, modelInfo.ModelName.Length - 4) + "\"
            End If
            storageDir = IO.Path.GetFullPath(storageDir)
            Return storageDir
        End Function

        Private Shared Function GetExcelTemplateName() As String
            Dim wbname = "component_"
            Dim JointNumber = getJointNumber()
            If JointNumber > 0 Then
                wbname = wbname + CStr(getJointNumber()) + ".xls"
            Else
                wbname = wbname + GetComponentName() + ".xls"
            End If
            Console.WriteLine(wbname)
            Return wbname
        End Function

        Shared Sub Run(context As CDelegate)
            Dim xlApp As Excel.Application
            Dim xlBooks As Excel.Workbooks
            Dim xlBook As Excel.Workbook
            Dim xlWorksheets As Excel.Sheets
            Dim xlSheetInputs As Excel.Worksheet
            Dim xlSheetOutputs As Excel.Worksheet
            Dim xlSheetComponent As Excel.Worksheet

            Dim inAttribute As Excel.Range
            Dim inValue As Excel.Range
            Dim inType As Excel.Range

            Dim outAttribute As Excel.Range
            Dim outValue As Excel.Range
            Dim outType As Excel.Range

            Dim jointStructAttribute As Excel.Range
            Dim jointStructValue As Excel.Range
            Dim jointAttribute As Excel.Range
            Dim jointValue As Excel.Range

            Iface = context

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            Dim doLoop As Boolean
            Dim pos As Integer
            Dim empty As Integer
            Dim basename As String
            Dim storageDir = GetOutputDirectory(Iface)
            Dim workbookName = GetExcelTemplateName()
            Dim workBookDir = GetWorkBookDirectory(workbookName, storageDir)
            If workBookDir.Equals("") Then
                Return
            End If
            
            Dim guid As String
			guid = getJointStructVariable("id")
            Dim resultname = storageDir + guid + "_res.xls"

            'Set culture to 'en-US'
            Dim oldCI As System.Globalization.CultureInfo = _
                System.Threading.Thread.CurrentThread.CurrentCulture
            System.Threading.Thread.CurrentThread.CurrentCulture = _
                New System.Globalization.CultureInfo("en-US")

            xlApp = CType(CreateObject("Excel.Application"), Excel.Application)
            xlBooks = CType(xlApp.Workbooks, Excel.Workbooks)

            Try
                xlBook = xlBooks.Open(workBookDir & workbookName)
            Catch NoSuchFile As System.Runtime.InteropServices.COMException
                    xlApp.Quit()
                    NAR(xlBooks)
                    NAR(xlApp)
                    Return
            End Try

            xlWorksheets = xlBook.Worksheets

            If SHOW_EXCEL Then
                xlApp.Visible = True
            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            Try
                xlSheetInputs = xlWorksheets("Inputs")
            Catch
                xlBook.Close(SaveChanges:=False)
                xlApp.Quit()
                Return
            End Try

            Try
                xlSheetOutputs = xlWorksheets("Outputs")
            Catch
                xlBook.Close(SaveChanges:=False)
                xlApp.Quit()
                Return
            End Try

            Try
                xlSheetComponent = xlWorksheets("Component")
            Catch
                xlBook.Close(SaveChanges:=False)
                xlApp.Quit()
                Return
            End Try


            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim xlBookNames As Excel.Names = xlBook.Names

            For idx As Integer = 1 To xlBookNames.Count
                Dim name As Excel.Name
                name = CType(xlBookNames.Item(idx), Excel.Name)
                'Console.WriteLine(name.Name)
                'If name.RefersToRange.Worksheet Is xlSheetInputs Then
                Dim nameName As String = name.Name
                If nameName.Equals("in_attribute") Then
                    inAttribute = name.RefersToRange
                ElseIf nameName.Equals("in_value") Then
                    inValue = name.RefersToRange
                ElseIf nameName.Equals("in_type") Then
                    inType = name.RefersToRange
                End If
                'End If

                'If name.RefersToRange.Worksheet Is xlSheetOutputs Then
                If nameName.Equals("out_attribute") Then
                    outAttribute = name.RefersToRange
                ElseIf nameName.Equals("out_value") Then
                    outValue = name.RefersToRange
                ElseIf nameName.Equals("out_type") Then
                    outType = name.RefersToRange
                End If
                'End If

                'If name.RefersToRange.Worksheet Is xlSheetComponent Then
                If nameName.Equals("jointstruct_attribute") Then
                    jointStructAttribute = name.RefersToRange
                ElseIf nameName.Equals("jointstruct_value") Then
                    jointStructValue = name.RefersToRange
                ElseIf nameName.Equals("joint_attribute") Then
                    jointAttribute = name.RefersToRange
                ElseIf nameName.Equals("joint_value") Then
                    jointValue = name.RefersToRange
                End If
                'End If
                NAR(name)
            Next


            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            xlSheetInputs.Activate()

            pos = 1
            empty = 0
            doLoop = True
            Do While doLoop
                pos += 1
                Dim attr As Excel.Range = inAttribute.Range("A" & pos)
                Dim value As Excel.Range = inValue.Range("A" & pos)
                Dim type As Excel.Range = inType.Range("A" & pos)
                If attr.Value Is Nothing Then
                    empty += 1
                    If empty > 10 Then
                        attr.Value = "END"
                        doLoop = False
                    End If
                ElseIf "END".Equals(attr.Value) Then
                    doLoop = False
                Else
                    empty = 0
                    Dim typelow As String

                    If type.Value Is Nothing Then
                        typelow = ""
                    Else
                        typelow = type.Value.ToLower()
                    End If

                    If "int".Equals(typelow) Then
                        Dim IntValue As Integer
                        If GetIntValue(attr.Value, IntValue) Then
                            value.Value = IntValue
                        End If
                    ElseIf "double".Equals(typelow) Then
                        Dim DblValue As Double
                        If GetDoubleValue(attr.Value, DblValue) Then
                            value.Value = DblValue
                        End If
                    ElseIf "string".Equals(typelow) Then
                        Dim StrValue As String
                        If GetStringValue(attr.Value, StrValue) Then
                            value.Value = StrValue
                        End If
                    End If
                End If
                NAR(attr)
                NAR(value)
                NAR(type)
            Loop

            If DEBUG Then
                Forms.MessageBox.Show("Input values are set.")
            End If

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            xlSheetComponent.Activate()

            Dim Joint As dotJoint_t
            GetJointStructure(Joint)


            pos = 1
            empty = 0
            doLoop = True
            Do While doLoop
                pos += 1
                Dim attr As Excel.Range = jointStructAttribute.Range("A" & pos)
                Dim value As Excel.Range = jointStructValue.Range("A" & pos)

                If attr.Value Is Nothing Then
                    empty += 1
                    If empty > 10 Then
                        attr.Value = "END"
                        doLoop = False
                    End If
                ElseIf "END".Equals(attr.Value) Then
                    doLoop = False
                Else
                    value.Value = getJointStructVariable(attr.Value)
                End If

                NAR(attr)
                NAR(value)
            Loop



            For xposi As Integer = 0 To getJointSecondariesCount()
                pos = 1
                empty = 0
                doLoop = True
                Dim xpos = Chr(Asc("A") + xposi)
                Do While doLoop
                    pos += 1
                    Dim attr As Excel.Range = jointAttribute.Range("A" & pos)
                    Dim value As Excel.Range = jointValue.Range(xpos & pos)

                    If attr.Value Is Nothing Then
                        empty += 1
                        If empty > 10 Then
                            attr.Value = "END"
                            doLoop = False
                        End If
                    ElseIf "END".Equals(attr.Value) Then
                        doLoop = False
                    Else
                        value.Value = getJointVariable(xposi, attr.Value)
                    End If
                    NAR(attr)
                    NAR(value)
                Loop
            Next

            If DEBUG Then
                Forms.MessageBox.Show("Component data set.")
            End If

            If RUN_MACRO Then
                xlApp.Run("StartCalculation")
            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            xlSheetOutputs.Activate()

            Dim Joint As dotJoint_t
            GetJointStructure(Joint)

            pos = 1
            empty = 0
            doLoop = True
            Do While doLoop
                pos += 1
                Dim xlRangeAttr As Excel.Range
                Dim xlRangeValue As Excel.Range
                Dim xlRangeType As Excel.Range
                Dim xlRangeAttrValue As Object
                Dim xlRangeValueValue As Object
                Dim xlRangeTypeValue As Object

                xlRangeAttr = outAttribute.Range("A" & pos)
                xlRangeValue = outValue.Range("A" & pos)
                xlRangeType = outType.Range("A" & pos)
                xlRangeAttrValue = xlRangeAttr.Value
                xlRangeValueValue = xlRangeValue.Value
                xlRangeTypeValue = xlRangeType.Value

                If xlRangeAttrValue Is Nothing Then
                    empty += 1
                    If empty > 10 Then
                        'xlRangeAttrValue = "END"
                        xlRangeAttrValue = "END.COMPLETED....."
                        doLoop = False
                    End If
                ElseIf "ENDEND.COMPLETED.....".Equals(xlRangeAttrValue) Then
                    doLoop = False
                ElseIf xlRangeAttrValue.ToString.StartsWith("UDA.") Then
                    empty = 0
                    Dim Name As String
                    Dim IdCount As Integer
                    Name = xlRangeAttrValue.ToString
                    Name = Name.Remove(0, 4)
                    If Name.StartsWith("10.") Then
                        IdCount = 10
                        Name.Remove(0, 3)
                    ElseIf Name.StartsWith("COM.") Then
                        IdCount = -1
                        Name = Name.Remove(0, 4)
                    Else
                        If Integer.TryParse(Name.Substring(0, 1), IdCount) Then
                            Name = Name.Remove(0, 2)
                        Else
                            Continue Do
                        End If
                    End If

                    If IdCount > Joint.nSecondaries Then
                        Continue Do
                    End If

                    Dim Id As Integer

                    If IdCount = -1 Then
                        Id = Joint.Id
                    ElseIf IdCount = 0 Then
                        Id = Joint.Primary
                    Else
                        Id = Joint.Secondaries(IdCount - 1)
                    End If

                    If "int".Equals(xlRangeTypeValue) Then
                        SetIntegerUdaValue(Id, Name, xlRangeValueValue)
                    ElseIf "string".Equals(xlRangeTypeValue) Then
                        SetStringUdaValue(Id, Name, xlRangeValueValue)
                    ElseIf "double".Equals(xlRangeTypeValue) Then
                        SetDoubleUdaValue(Id, Name, xlRangeValueValue)
                    End If
                ElseIf "int".Equals(xlRangeTypeValue) Then
                    SetIntValue(xlRangeAttrValue, xlRangeValueValue)
                ElseIf "double".Equals(xlRangeTypeValue) Then
                    SetDoubleValue(xlRangeAttrValue, xlRangeValueValue)
                ElseIf "string".Equals(xlRangeTypeValue) Then
                    SetStringValue(xlRangeAttrValue, xlRangeValueValue)
                End If

                NAR(xlRangeAttr)
                NAR(xlRangeValue)
                NAR(xlRangeType)
            Loop

            If DEBUG Then
                Forms.MessageBox.Show("Output values are read.")
            End If

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            ' Hide all sheets except "Joint"
            Dim xlWorksheetsCount As Integer = xlWorksheets.Count
            Dim xlWorksheet As Excel.Worksheet
            For idx As Integer = 1 To xlWorksheetsCount
                xlWorksheet = xlWorksheets(idx)
                If Not xlWorksheet Is xlSheetOutputs Then
                    xlWorksheet.Visible = False
                End If
                NAR(xlWorksheet)
            Next
            'xlSheetOutputs.Visible = True
            'xlSheetInputs.Visible = False
            'xlSheetComponent.Visible = False

            If STORE_RESULTS Then
                Try
                    If Not IO.Directory.Exists(storageDir) Then
                        IO.Directory.CreateDirectory(storageDir)
                    End If
                    If IO.File.Exists(resultname) Then
                        IO.File.Delete(resultname)
                    End If
                Catch ex As Exception
                    If DEBUG Then
                        Forms.MessageBox.Show("Could not remove file '" & resultname & "'." & ex.Message)
                    End If
                End Try
                Try
                    xlSheetOutputs.SaveAs(resultname)
                Catch ex As Exception
                    Forms.MessageBox.Show("Could not save file '" & resultname & "'." & ex.Message)
                End Try
            End If

            If False Then 'XXX
            End If 'XXX

            NAR(xlBookNames)

            NAR(inAttribute)
            NAR(inValue)
            NAR(inType)
            NAR(outAttribute)
            NAR(outValue)
            NAR(outType)
            NAR(jointStructAttribute)
            NAR(jointStructValue)
            NAR(jointAttribute)
            NAR(jointValue)

            NAR(xlSheetComponent)
            NAR(xlSheetOutputs)
            NAR(xlSheetInputs)
            NAR(xlWorksheets)

            xlBook.Close(SaveChanges:=False)
            NAR(xlBook)
            NAR(xlBooks)
            xlApp.Quit()
            NAR(xlApp)

            GC.Collect()
            System.Threading.Thread.CurrentThread.CurrentCulture = oldCI

        End Sub
    End Class
End Namespace
