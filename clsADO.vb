Imports System.Data
Imports System.Data.OracleClient
'Imports Oracle.DataAccess.Client
'Imports Oracle.DataAccess.Types
Imports ADODB

Public Class clsADO

    Dim Conn_SQL As New SqlClient.SqlConnection
    Dim Conn_ADO As New ADODB.Connection
    'Dim OracleConn As New System.Data.OracleClient.OracleConnection
    Dim conn_ORACLE As New OracleConnection

    '資料庫
    Public ServerName As String
    Public Password As String
    Public LoginName As String
    Public DBname As String
    Public connectOK As Boolean
    Public glConnStringNet As String
    Public glConnStringADO As String
    Public DBStyle As String
    Public blnWindowsAuthen As Boolean '是否使用window 驗證

    'Public AccessConnDB As New OleDbConnection
    'Public OracleConn As New OracleClient.OracleConnection



    Public ReadOnly Property SelConnection As ADODB.Connection
        Get
            If DBStyle = "SQL" Then
                Return Conn_ADO
            Else
                Return Conn_ADO
            End If

        End Get
    End Property

    Public ReadOnly Property SelConnectionORACLE As OracleConnection
        Get
            'If DBStyle = "SQL" Then
            '    Return Conn_SQL
            'Else
            Return conn_ORACLE
            'End If

        End Get
    End Property

    Public ReadOnly Property SelConnectionSQL As SqlClient.SqlConnection
        Get
            'If DBStyle = "SQL" Then
            Return Conn_SQL
            'Else
            'Return conn_ORACLE
            'End If

        End Get
    End Property

    Public Sub InitialtionDB()
        If InStr(gcDBString, "Provider=MSDAORA") Then
            'ORACLE
            DBStyle = "ORACLE"
            glConnStringNet = "Data Source=" & ServerName & ";Persist Security Info=True;User ID=" & LoginName & ";Password=" & Password & ";Unicode=True"
            glConnStringADO = "Provider=MSDAORA.1;Password='" & Password & "';User ID=" & LoginName & ";Data Source=" & ServerName & ";Persist Security Info=True"
        Else
            'SQL Server
            DBStyle = "SQL"
            If blnWindowsAuthen Then
                glConnStringNet = "Data Source=" & ServerName & ";Initial Catalog=" & DBname & ";Integrated Security=True;packet size=4096;Connection Timeout=45"
                glConnStringADO = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=sa;Initial Catalog=" & DBname & ";Data Source=" & ServerName & ";Use Procedure for Prepare=1;packet size=4096;Connection Timeout=45"
            Else
                glConnStringNet = "data source=" & ServerName & ";initial catalog=" & DBname & ";password='" & Password & "';persist security info=True;user id=" & LoginName & ";packet size=4096;Connection Timeout=45"
                glConnStringADO = "Provider=SQLOLEDB.1;Password='" & Password & "';Persist Security Info=True;User ID=" & LoginName & ";Initial Catalog=" & DBname & ";Data Source=" & ServerName
            End If

        End If
 
    End Sub

    Public Function OpenConnection() As Boolean
        Try

            'ADO
            If Conn_ADO.State <> ADODB.ObjectStateEnum.adStateClosed Then Conn_ADO.Close()
            Conn_ADO.ConnectionString = glConnStringADO
            Conn_ADO.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            Conn_ADO.ConnectionTimeout = 45
            Conn_ADO.Open()

            'Dot NET 
            If glAdoInst.DBStyle = "ORACLE" Then
                If conn_ORACLE.State <> ConnectionState.Closed Then conn_ORACLE.Close()
                conn_ORACLE.ConnectionString = glConnStringNet
                conn_ORACLE.Open()
                connectOK = True
            Else
                If Conn_SQL.State <> ConnectionState.Closed Then Conn_SQL.Close()
                Conn_SQL.ConnectionString = glConnStringNet
                Conn_SQL.Open()
                connectOK = True
            End If
 
            CheckDBVer()

            Return True
        Catch ex As Exception
            MsgBox(Err.Description & vbNewLine & "Please reset the Database!", MsgBoxStyle.Critical)
            connectOK = False
            Return False
            'End
        End Try
    End Function
    'Public Function GetNewDT(ByVal cSQL As String) As DataTable

    '    Dim da As SqlClient.SqlDataAdapter
    '    Dim DT As New DataTable
    '    'Select Case gcDBType
    '    '    Case "S"
    '    da = New SqlClient.SqlDataAdapter(cSQL, ConnDB)
    '    da.Fill(DT)
    '    '    Case "C"
    '    'daC = New OleDbDataAdapter(cSQL, AccessConnDB)
    '    'daC.Fill(DT)
    '    '    Case "O"
    '    'daO = New OracleClient.OracleDataAdapter(cSQL, OracleConn)
    '    'daO.Fill(DT)

    '    'End Select
    '    Return DT

    'End Function
    Public Function CreatRecordset(ByVal cSQL As String, ByVal connStr As String, Optional ByVal CursorType As CursorTypeEnum = CursorTypeEnum.adOpenDynamic, Optional ByVal LockType As LockTypeEnum = LockTypeEnum.adLockBatchOptimistic) As Recordset
        Dim rs As New ADODB.Recordset
        Try

            'Select Case DBStyle
            'Case "SQL"
            rs.Open(cSQL, Conn_ADO, CursorType, LockType)
            'Case "ORACLE"
            'End Select


        Catch ex As Exception

        End Try
        Return rs

    End Function

    Public ReadOnly Property GetNewDT(ByVal cSQL As String, Optional ByVal myTrans As Object = Nothing) As DataTable
        Get
            Dim DT As New DataTable
            If glAdoInst.DBStyle = "ORACLE" Then
                Dim da As OracleDataAdapter
                Dim cb As OracleCommandBuilder
                da = New OracleDataAdapter(cSQL, conn_ORACLE)
                cb = New OracleCommandBuilder(da)
                If Not IsNothing(myTrans) Then
                    If Not IsNothing(da.InsertCommand) Then
                        da.InsertCommand.Transaction = myTrans
                    End If
                    If Not IsNothing(da.UpdateCommand) Then
                        da.UpdateCommand.Transaction = myTrans
                    End If
                    If Not IsNothing(da.DeleteCommand) Then
                        da.DeleteCommand.Transaction = myTrans
                    End If
                    If Not IsNothing(da.SelectCommand) Then
                        da.SelectCommand.Transaction = myTrans
                    End If
                End If
                da.Fill(DT)
                da.Dispose()
                cb.Dispose()
            Else
                Dim da As SqlClient.SqlDataAdapter
                Dim cb As SqlClient.SqlCommandBuilder
                da = New SqlClient.SqlDataAdapter(cSQL, Conn_SQL)
                cb = New SqlClient.SqlCommandBuilder(da)

                If Not IsNothing(myTrans) Then
                    If Not IsNothing(da.InsertCommand) Then
                        da.InsertCommand.Transaction = myTrans
                    End If
                    If Not IsNothing(da.UpdateCommand) Then
                        da.UpdateCommand.Transaction = myTrans
                    End If
                    If Not IsNothing(da.DeleteCommand) Then
                        da.DeleteCommand.Transaction = myTrans
                    End If
                    If Not IsNothing(da.SelectCommand) Then
                        da.SelectCommand.Transaction = myTrans
                    End If
                End If

                da.Fill(DT)
                da.Dispose()
                cb.Dispose()
            End If

            Return DT
        End Get
    End Property

    Public Sub ExcuteSQL(ByVal cSQL As String, ByRef myTrans As Object, Optional ByVal blnThrowException As Boolean = False)
        '只針對 MQ_ExFiles update

        Try
            If glAdoInst.DBStyle = "ORACLE" Then
                Dim cmd As New OracleClient.OracleCommand(cSQL, SelConnectionORACLE, myTrans)

                cmd.ExecuteNonQuery()

                cmd.Dispose()


            Else

                Dim cmd As New SqlClient.SqlCommand(cSQL, SelConnectionSQL, myTrans)

                cmd.ExecuteNonQuery()

                cmd.Dispose()


                'Dim da As SqlClient.SqlDataAdapter
                'Dim cb As SqlClient.SqlCommandBuilder
                'da = New SqlClient.SqlDataAdapter(cSQL, glAdoInst.SelConnectionSQL)
                'cb = New SqlClient.SqlCommandBuilder(da)

                'If Not IsNothing(myTrans) Then
                '    If Not IsNothing(da.InsertCommand) Then
                '        da.InsertCommand.Transaction = myTrans
                '    End If
                '    If Not IsNothing(da.UpdateCommand) Then
                '        da.UpdateCommand.Transaction = myTrans
                '    End If
                '    If Not IsNothing(da.DeleteCommand) Then
                '        da.DeleteCommand.Transaction = myTrans
                '    End If
                '    If Not IsNothing(da.SelectCommand) Then
                '        da.SelectCommand.Transaction = myTrans
                '    End If
                'End If

                'da.UpdateCommand.ExecuteNonQuery()

                'da.Dispose()
                'cb.Dispose()


            End If

        Catch exC As System.Data.DBConcurrencyException

            MsgBox("資料同步化錯誤，可能是網路上有其他使用者同時更改資料!")

            If IsNothing(myTrans) Then
                MsgBox(exC.ToString)
            Else
                Throw exC
            End If
            'Catch ex As System.Data.SqlClient.SqlException
            '    Throw ex
        Catch ex As Exception
            If blnThrowException Then
                Throw ex
            ElseIf IsNothing(myTrans) Then
                MsgBox(ex.ToString)

            Else
                Throw ex
            End If
        End Try
    End Sub

    Public Sub UpdateDT(ByVal cSQL As String, ByRef dtUpdate As DataTable, Optional ByRef myTrans As Object = Nothing, Optional ByVal blnThrowException As Boolean = False)
        '只針對 MQ_ExFiles update

        Try
            If glAdoInst.DBStyle = "ORACLE" Then
                Dim da As OracleDataAdapter
                Dim cb As OracleCommandBuilder
                da = New OracleDataAdapter(cSQL, glAdoInst.SelConnectionORACLE)
                cb = New OracleCommandBuilder(da)

                If Not IsNothing(myTrans) Then
                    If Not IsNothing(da.InsertCommand) Then
                        da.InsertCommand.Transaction = myTrans
                    End If
                    If Not IsNothing(da.UpdateCommand) Then
                        da.UpdateCommand.Transaction = myTrans
                    End If
                    If Not IsNothing(da.DeleteCommand) Then
                        da.DeleteCommand.Transaction = myTrans
                    End If
                    If Not IsNothing(da.SelectCommand) Then
                        da.SelectCommand.Transaction = myTrans
                    End If
                End If

                da.Update(dtUpdate)
                da.Dispose()
                cb.Dispose()

            Else
                Dim da As SqlClient.SqlDataAdapter
                Dim cb As SqlClient.SqlCommandBuilder
                da = New SqlClient.SqlDataAdapter(cSQL, glAdoInst.SelConnectionSQL)
                cb = New SqlClient.SqlCommandBuilder(da)

                If Not IsNothing(myTrans) Then
                    If Not IsNothing(da.InsertCommand) Then
                        da.InsertCommand.Transaction = myTrans
                    End If
                    If Not IsNothing(da.UpdateCommand) Then
                        da.UpdateCommand.Transaction = myTrans
                    End If
                    If Not IsNothing(da.DeleteCommand) Then
                        da.DeleteCommand.Transaction = myTrans
                    End If
                    If Not IsNothing(da.SelectCommand) Then
                        da.SelectCommand.Transaction = myTrans
                    End If
                End If

                da.Update(dtUpdate)
                da.Dispose()
                cb.Dispose()


            End If

        Catch exC As System.Data.DBConcurrencyException

            MsgBox("資料同步化錯誤，可能是網路上有其他使用者同時更改資料!")

            If IsNothing(myTrans) Then
                MsgBox(exC.ToString)
            Else
                Throw exC
            End If
            'Catch ex As System.Data.SqlClient.SqlException
            '    Throw ex
        Catch ex As Exception
            If blnThrowException Then
                Throw ex
            ElseIf IsNothing(myTrans) Then
                MsgBox(ex.ToString)

            Else
                Throw ex
            End If
        End Try
    End Sub

    '物件導向之多型  dtUpdate ->  datarow
    Public Sub UpdateDT(ByVal cSQL As String, ByRef drUpdate() As DataRow, Optional ByRef myTrans As Object = Nothing, Optional ByVal blnThrowException As Boolean = False)
        '只針對 MQ_ExFiles update

        Try
            If glAdoInst.DBStyle = "ORACLE" Then
                Dim da As OracleDataAdapter
                Dim cb As OracleCommandBuilder
                da = New OracleDataAdapter(cSQL, glAdoInst.SelConnectionORACLE)
                cb = New OracleCommandBuilder(da)

                If Not IsNothing(myTrans) Then
                    If Not IsNothing(da.InsertCommand) Then
                        da.InsertCommand.Transaction = myTrans
                    End If
                    If Not IsNothing(da.UpdateCommand) Then
                        da.UpdateCommand.Transaction = myTrans
                    End If
                    If Not IsNothing(da.DeleteCommand) Then
                        da.DeleteCommand.Transaction = myTrans
                    End If
                    If Not IsNothing(da.SelectCommand) Then
                        da.SelectCommand.Transaction = myTrans
                    End If
                End If

                da.Update(drUpdate)
                da.Dispose()
                cb.Dispose()

            Else
                Dim da As SqlClient.SqlDataAdapter
                Dim cb As SqlClient.SqlCommandBuilder
                da = New SqlClient.SqlDataAdapter(cSQL, glAdoInst.SelConnectionSQL)
                cb = New SqlClient.SqlCommandBuilder(da)

                If Not IsNothing(myTrans) Then
                    If Not IsNothing(da.InsertCommand) Then
                        da.InsertCommand.Transaction = myTrans
                    End If
                    If Not IsNothing(da.UpdateCommand) Then
                        da.UpdateCommand.Transaction = myTrans
                    End If
                    If Not IsNothing(da.DeleteCommand) Then
                        da.DeleteCommand.Transaction = myTrans
                    End If
                    If Not IsNothing(da.SelectCommand) Then
                        da.SelectCommand.Transaction = myTrans
                    End If
                End If

                da.Update(drUpdate)
                da.Dispose()
                cb.Dispose()


            End If

        Catch exC As System.Data.DBConcurrencyException

            MsgBox("資料同步化錯誤，可能是網路上有其他使用者同時更改資料!")

            If IsNothing(myTrans) Then
                MsgBox(exC.ToString)
            Else
                Throw exC
            End If
            'Catch ex As System.Data.SqlClient.SqlException
            '    Throw ex
        Catch ex As Exception
            If blnThrowException Then
                Throw ex
            ElseIf IsNothing(myTrans) Then
                MsgBox(ex.ToString)

            Else
                Throw ex
            End If
        End Try
    End Sub





    Public Function GetAutoNumber(ByVal mField As String) As Long
        Dim cID As Long
        Dim GetAgain As Boolean
        Dim rsAutoNumber As New Recordset


        'On Error GoTo ErrGetID
        Try
            If Conn_ADO.State <> ADODB.ObjectStateEnum.adStateOpen Then Conn_ADO.Open()

            rsAutoNumber.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            rsAutoNumber.Open("Select * From MQ_AutoNumber", Conn_ADO, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

            If rsAutoNumber.RecordCount = 0 Then
                Conn_ADO.Execute("Insert Into MQ_AutoNumber (LIGID,MQFileID,ShipmentContrast) Values (100,100,100)")
            End If
        Catch ex As Exception
            MsgBox("GetAutoNumber_1" & vbCrLf & ex.ToString)
        End Try

        'transaction
        Try
            Conn_ADO.BeginTrans()
            '取得自動編號
            GetAgain = True
            Do While GetAgain
                rsAutoNumber.Requery()
                rsAutoNumber.MoveFirst()
                If IsDBNull(rsAutoNumber.Fields(mField).Value) Then
                    rsAutoNumber.Fields(mField).Value = 1000
                    cID = 1000
                Else
                    cID = rsAutoNumber.Fields(mField).Value
                End If

                rsAutoNumber.Fields(mField).Value = cID + 1
                GetAgain = False
                rsAutoNumber.Update()
            Loop

            GetAutoNumber = cID

            rsAutoNumber.Close()
            rsAutoNumber = Nothing
            Conn_ADO.CommitTrans()
        Catch ex As Exception
            Conn_ADO.RollbackTrans()
            GetAutoNumber = 0
            MsgBox("GetAutoNumber_2" & vbCrLf & ex.ToString)

        End Try

    End Function
    Private Sub CheckDBVer()
        Dim rs As New ADODB.Recordset
        Dim cSQL As String
        Dim ans As MsgBoxResult
        'initial DB
        Try
            rs.Open("Select * from MQ_Batch", glAdoInst.SelConnection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        Catch ex As Exception
            ans = MsgBox("使用 MQFC 需要升級資料庫，請問是否升級資料庫?", MsgBoxStyle.YesNo)

            If ans = MsgBoxResult.Yes Then

                Try
                    glAdoInst.SelConnection.BeginTrans()
                    cSQL = "CREATE TABLE [dbo].[MQ_AutoNumber](" & _
                    "[LIGID] [bigint] NOT NULL," & _
                    "[MQFileID] [bigint] NULL," & _
                    "[ShipmentContrast] [bigint] NULL" & _
                    ") ON [PRIMARY]"
                    glAdoInst.SelConnection.Execute(cSQL)


                    cSQL = "CREATE TABLE [dbo].[MQ_Batch](" & _
                         "[MQBatchID] [bigint] IDENTITY(1,1) NOT NULL," & _
                         "[MQBatchNo] [nvarchar](150)  NULL," & _
                         "[MQFileID] [bigint] NOT NULL," & _
                         "[ExpectOutQty] [float] NULL," & _
                         "[LevelGroup] [nvarchar](100) NULL," & _
                         "[CreateDate] [datetime] NULL," & _
                         "[Finished] [smallint] NULL," & _
                         "[FinishDate] [datetime] NULL," & _
                         "[Level1] [nvarchar](100) NULL," & _
                         "[Level2] [nvarchar](100) NULL," & _
                         "[Level3] [nvarchar](100) NULL," & _
                         "[Level4] [nvarchar](100) NULL," & _
                         "[Level5] [nvarchar](100) NULL," & _
                         "[Level6] [nvarchar](100) NULL," & _
                         "[Level7] [nvarchar](100) NULL," & _
                         "[Level8] [nvarchar](100) NULL," & _
                         "[Level9] [nvarchar](100) NULL," & _
                         "[Level10] [nvarchar](100) NULL," & _
                         "[MQResBatch1] [nvarchar](100) NULL," & _
                         "[MQResBatch2] [nvarchar](100) NULL," & _
                         "[MQResBatch3] [nvarchar](100) NULL," & _
                         "CONSTRAINT [PK_MQ_Batch] PRIMARY KEY CLUSTERED " & _
                         "([MQBatchID] Asc" & _
                         ")WITH (PAD_INDEX  = OFF, IGNORE_DUP_KEY = OFF) ON [PRIMARY]" & _
                         ") ON [PRIMARY]"
                    glAdoInst.SelConnection.Execute(cSQL)


                    cSQL = "CREATE TABLE [dbo].[MQ_CheckIn](" & _
                         "[MQCheckIn_ID] [bigint] IDENTITY(1,1) NOT NULL," & _
                         "[MaterialNo] [nvarchar](150) NOT NULL," & _
                         "[MQFileID] [bigint] NOT NULL," & _
                         "[MQBatchID] [bigint] NOT NULL," & _
                         "[CheckInItem] [smallint] NULL," & _
                         "[CheckInQty] [float] NULL," & _
                         "[LeftQty] [float] NULL," & _
                         "[UserDefine1] [nvarchar](100) NULL," & _
                         "[UserDefine2] [nvarchar](100) NULL," & _
                         "[UserDefine3] [nvarchar](100) NULL," & _
                         "[UserDefine4] [nvarchar](100) NULL," & _
                         "[UserDefine5] [nvarchar](100) NULL," & _
                         "[CheckInDate] [datetime] NULL," & _
                         "CONSTRAINT [PK_MQ_CheckIn] PRIMARY KEY CLUSTERED " & _
                         "(" & _
                         "[MQCheckIn_ID] Asc" & _
                         ")WITH (PAD_INDEX  = OFF, IGNORE_DUP_KEY = OFF) ON [PRIMARY]" & _
                         ") ON [PRIMARY]"
                    glAdoInst.SelConnection.Execute(cSQL)


                    cSQL = "CREATE TABLE [dbo].[MQ_CheckOut](" & _
                             "[CheckOutID] [bigint] IDENTITY(1,1) NOT NULL," & _
                             "[MaterialNo] [nvarchar](150) NOT NULL," & _
                             "[MQFileID] [bigint] NOT NULL," & _
                             "[MQBatchID] [bigint] NOT NULL," & _
                             "[EXMQBatchID] [bigint] NULL," & _
                             "[MQCheckIn_ID] [smallint] NULL," & _
                             "[GetItem] [smallint] NULL," & _
                             "[GetQty] [float] NULL," & _
                             "[ResMQRev1] [nvarchar](100) NULL," & _
                             "[ResMQRev2] [nvarchar](100) NULL," & _
                             "[ResMQRev3] [nvarchar](100) NULL," & _
                             "CONSTRAINT [PK_MQ_BatchRelation] PRIMARY KEY CLUSTERED " & _
                            "(" & _
                            "[CheckOutID] Asc" & _
                            ")WITH (PAD_INDEX  = OFF, IGNORE_DUP_KEY = OFF) ON [PRIMARY]" & _
                            ") ON [PRIMARY]"
                    glAdoInst.SelConnection.Execute(cSQL)


                    cSQL = "CREATE TABLE [dbo].[MQ_EXFiles](" & _
                             "[PreID] [bigint] IDENTITY(1,1) NOT NULL," & _
                             "[MQFileID] [bigint] NOT NULL," & _
                             "[ExFileID] [bigint] NOT NULL," & _
                             "[InOutRate] [float] NULL," & _
                             "CONSTRAINT [PK_MQ_FileRelation] PRIMARY KEY CLUSTERED " & _
                            "(" & _
                                        "[PreID] Asc" & _
                            ")WITH (PAD_INDEX  = OFF, IGNORE_DUP_KEY = OFF) ON [PRIMARY]" & _
                            ") ON [PRIMARY]"
                    glAdoInst.SelConnection.Execute(cSQL)


                    cSQL = "CREATE TABLE [dbo].[MQ_FileLevel](" & _
                         "[MQFileID] [bigint] NOT NULL," & _
                         "[LevelID] [int] NOT NULL," & _
                         "[LevelName] [nvarchar](100) NULL," & _
                         "[LevelItem] [nvarchar](100) NULL," & _
                         "CONSTRAINT [PK_MQ_FileLevel] PRIMARY KEY CLUSTERED " & _
                        "(" & _
                        "[MQFileID] ASC," & _
                        "[LevelID] Asc" & _
                        ")WITH (PAD_INDEX  = OFF, IGNORE_DUP_KEY = OFF) ON [PRIMARY]" & _
                        ") ON [PRIMARY]"
                    glAdoInst.SelConnection.Execute(cSQL)


                    cSQL = "CREATE TABLE [dbo].[MQ_FileMaterial](" & _
                         "[FMID] [bigint] IDENTITY(1,1) NOT NULL," & _
                         "[MQFileID] [bigint] NOT NULL," & _
                         "[MaterialNo] [nvarchar](150) NOT NULL," & _
                         "[IsRate] [smallint] NOT NULL," & _
                         "[InOutRate] [float] NOT NULL," & _
                         "[LossRate] [float] NOT NULL," & _
                         "[Reserve1] [nvarchar](50) NULL," & _
                         "[Reserve2] [nvarchar](50) NULL," & _
                         "CONSTRAINT [PK_MQ_FileMaterial] PRIMARY KEY CLUSTERED " & _
                        "(" & _
                                    "[FMID] Asc" & _
                        ")WITH (PAD_INDEX  = OFF, IGNORE_DUP_KEY = OFF) ON [PRIMARY]" & _
                        ") ON [PRIMARY]"
                    glAdoInst.SelConnection.Execute(cSQL)


                    cSQL = "CREATE TABLE [dbo].[MQ_FileName](" & _
                         "[MQFileID] [bigint] NOT NULL," & _
                         "[MaterialNo] [nvarchar](150) NULL," & _
                         "[MQFileName] [nvarchar](100) NOT NULL," & _
                         "[GroupName] [nvarchar](200) NULL," & _
                         "[CreateDate] [datetime] NULL," & _
                         "[LastChangeDate] [datetime] NULL," & _
                         "[LevelGroup] [bigint] NULL," & _
                         "[UserDefine1] [nvarchar](100) NULL," & _
                         "[UserDefine2] [nvarchar](100) NULL," & _
                         "[UserDefine3] [nvarchar](100) NULL," & _
                         "[UserDefine4] [nvarchar](100) NULL," & _
                         "[UserDefine5] [nvarchar](100) NULL," & _
                         "[UserDefine6] [nvarchar](100) NULL," & _
                         "[UserDefine7] [nvarchar](100) NULL," & _
                         "[UserDefine8] [nvarchar](100) NULL," & _
                         "[UserDefine9] [nvarchar](100) NULL," & _
                         "[UserDefine10] [nvarchar](100) NULL," & _
                         "[MQResFile1] [nvarchar](100) NULL," & _
                         "[MQResFile2] [nvarchar](100) NULL," & _
                         "[MQResFile3] [nvarchar](100) NULL," & _
                        " CONSTRAINT [PK_MQ_FileName] PRIMARY KEY CLUSTERED " & _
                        "(" & _
                                    "[MQFileID] Asc" & _
                        ")WITH (PAD_INDEX  = OFF, IGNORE_DUP_KEY = OFF) ON [PRIMARY]" & _
                        ") ON [PRIMARY]"
                    glAdoInst.SelConnection.Execute(cSQL)


                    cSQL = "CREATE TABLE [dbo].[MQ_GroupName](" & _
                             "[GID] [bigint] IDENTITY(1,1) NOT NULL," & _
                             "[GroupID] [nvarchar](50) NULL," & _
                             "[GroupName] [nvarchar](100) NOT NULL," & _
                             "CONSTRAINT [PK_MQ_GroupName] PRIMARY KEY CLUSTERED " & _
                            "(" & _
                                        "[GID] Asc" & _
                            ")WITH (PAD_INDEX  = OFF, IGNORE_DUP_KEY = OFF) ON [PRIMARY]" & _
                            ") ON [PRIMARY]"
                    glAdoInst.SelConnection.Execute(cSQL)


                    cSQL = "CREATE TABLE [dbo].[MQ_LEVELITEM](" & _
                             "[LIID] [bigint] IDENTITY(1,1) NOT NULL," & _
                             "[LEVELID] [smallint] NOT NULL," & _
                             "[EITEMNAME] [nvarchar](100) NOT NULL," & _
                             "[ITEMNAME] [nvarchar](100) NOT NULL," & _
                             "CONSTRAINT [PK_MQ_LEVELITEM] PRIMARY KEY CLUSTERED " & _
                            "(" & _
                             "[LEVELID] ASC," & _
                             "[EITEMNAME] ASC," & _
                                        "[ITEMNAME] Asc" & _
                            ")WITH (PAD_INDEX  = OFF, IGNORE_DUP_KEY = OFF) ON [PRIMARY]" & _
                            ") ON [PRIMARY]"
                    glAdoInst.SelConnection.Execute(cSQL)


                    cSQL = "CREATE TABLE [dbo].[MQ_LEVELITEMGROUP](" & _
                         "[LIGID] [bigint] NOT NULL," & _
                         "[ITEMGROUPNAME] [nvarchar](100) NOT NULL," & _
                         "CONSTRAINT [PK_MQ_LEVELITEMGROUP] PRIMARY KEY CLUSTERED " & _
                        "(" & _
                                    "[LIGID] Asc" & _
                        ")WITH (PAD_INDEX  = OFF, IGNORE_DUP_KEY = OFF) ON [PRIMARY]" & _
                        ") ON [PRIMARY]"
                    glAdoInst.SelConnection.Execute(cSQL)


                    cSQL = "CREATE TABLE [dbo].[MQ_LevelItemRelation](" & _
                             "[LIGID] [bigint] NOT NULL," & _
                             "[LIID] [bigint] NOT NULL," & _
                             "CONSTRAINT [PK_MQ_LevelItemRelation] PRIMARY KEY CLUSTERED " & _
                            "(" & _
                             "[LIGID] ASC," & _
                            "[LIID] Asc" & _
                            ")WITH (PAD_INDEX  = OFF, IGNORE_DUP_KEY = OFF) ON [PRIMARY]" & _
                            ") ON [PRIMARY]"
                    glAdoInst.SelConnection.Execute(cSQL)


                    cSQL = "CREATE TABLE [dbo].[MQ_LEVELNAME](" & _
                         "[LEVELID] [int] NOT NULL," & _
                         "[COLName] [nvarchar](50) NULL," & _
                         "[LEVELNAME] [nvarchar](100) NOT NULL," & _
                         "[ELEVELNAME] [nvarchar](100) NULL," & _
                         "[ACTIVED] [smallint] NOT NULL," & _
                         "[PosOrder] [smallint] NULL," & _
                         "CONSTRAINT [PK_LEVELNAME] PRIMARY KEY CLUSTERED " & _
                        "(" & _
                        "[LEVELID] Asc" & _
                        ")WITH (PAD_INDEX  = OFF, IGNORE_DUP_KEY = OFF) ON [PRIMARY]" & _
                        ") ON [PRIMARY]"
                    glAdoInst.SelConnection.Execute(cSQL)


                    cSQL = "CREATE TABLE [dbo].[MQ_Material](" & _
                             "[MaterialNo] [nvarchar](150) NOT NULL," & _
                             "[MaterialName] [nvarchar](150) NULL," & _
                             "[Provider] [nvarchar](100) NULL," & _
                             "[Memo] [nvarchar](100) NULL," & _
                             "[MaterialUser1] [nvarchar](100) NULL," & _
                             "[MaterialUser2] [nvarchar](100) NULL," & _
                             "[MaterialUser3] [nvarchar](100) NULL," & _
                             "CONSTRAINT [PK_MQ_Material] PRIMARY KEY CLUSTERED " & _
                            "(" & _
                            "[MaterialNo] Asc" & _
                            ")WITH (PAD_INDEX  = OFF, IGNORE_DUP_KEY = OFF) ON [PRIMARY]" & _
                            ") ON [PRIMARY]"

                    glAdoInst.SelConnection.Execute(cSQL)

                    cSQL = "CREATE TABLE [dbo].[MQ_QCFiles](" & _
                         "[QCFileID] [bigint] IDENTITY(1,1) NOT NULL," & _
                         "[MQFileID] [bigint] NOT NULL," & _
                         "[QCFileType] [smallint] NOT NULL," & _
                         "[FileID] [bigint] NOT NULL," & _
                         "CONSTRAINT [PK_MQ_QCFiles] PRIMARY KEY CLUSTERED " & _
                        "([QCFileID] Asc" & _
                        ")WITH (PAD_INDEX  = OFF, IGNORE_DUP_KEY = OFF) ON [PRIMARY]" & _
                        ") ON [PRIMARY]"
                    glAdoInst.SelConnection.Execute(cSQL)


                    cSQL = "CREATE TABLE [dbo].[MQ_QCLEVEL](" & _
                         "[QCLevelID] [int] NOT NULL," & _
                         "[QCLevelName] [nvarchar](100) NOT NULL," & _
                         "[QCLevelEnable] [smallint] NOT NULL," & _
                         "[QCLevelRev1] [nvarchar](50) NULL," & _
                         "[QCLevelRev2] [nvarchar](50) NULL," & _
                         "CONSTRAINT [PK_MQ_QCLEVEL] PRIMARY KEY CLUSTERED " & _
                        "(" & _
                        "[QCLevelID] Asc" & _
                        ")WITH (PAD_INDEX  = OFF, IGNORE_DUP_KEY = OFF) ON [PRIMARY]" & _
                        ") ON [PRIMARY]"
                    glAdoInst.SelConnection.Execute(cSQL)


                    cSQL = "CREATE TABLE [dbo].[MQ_UserDefine](" & _
                         "[UID] [bigint] IDENTITY(1,1) NOT NULL," & _
                         "[UserType] [nvarchar](50) NOT NULL," & _
                         "[UNo] [bigint] NOT NULL," & _
                         "[UpperNo] [bigint] NULL," & _
                         "[UserID] [nvarchar](50) NULL," & _
                         "[UserData] [nvarchar](100) NULL," & _
                         "[UserGroup] [smallint] NULL," & _
                         "[IsTopLevel] [smallint] NOT NULL," & _
                         "[Enabled] [smallint] NULL," & _
                         "[UserRev1] [nvarchar](50) NULL," & _
                         "[UserRev2] [nvarchar](50) NULL," & _
                         "CONSTRAINT [PK_MQ_UserDefine] PRIMARY KEY CLUSTERED " & _
                        "(" & _
                        "[UID] Asc" & _
                        ")WITH (PAD_INDEX  = OFF, IGNORE_DUP_KEY = OFF) ON [PRIMARY]" & _
                        ") ON [PRIMARY]"
                    glAdoInst.SelConnection.Execute(cSQL)

                    glAdoInst.SelConnection.CommitTrans()
                Catch ex2 As Exception
                    glAdoInst.SelConnection.RollbackTrans()
                    MsgBox("資料庫升級失敗，回復成原始狀態!")
                End Try




            End If
        End Try




        'ver1004
        Try
            If rs.State <> ADODB.ObjectStateEnum.adStateClosed Then rs.Close()
            rs.Open("Select * from MQ_DBVer", glAdoInst.SelConnection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        Catch ex As Exception


            glAdoInst.SelConnection.Execute("Create Table MQ_DBVer (ID [nvarchar](50), DBVer [nvarchar](50))")
            glAdoInst.SelConnection.Execute("Insert Into MQ_DBVer (ID,DBVer) values ('DBVer','1004')")
            glAdoInst.SelConnection.Execute( _
           "CREATE PROCEDURE dbo.InsertMQ_Batch  @MQFileID bigint,@Finished smallint,@CreateDate datetime, @Identity bigint OUT " & _
           "AS INSERT INTO MQ_Batch (MQFileID,ExpectOutQty,CreateDate,Finished,MQBatchNo,Level1,Level2,Level3,Level4,Level5,Level6,Level7,Level8,Level9,Level10) " & _
           "VALUES(@MQFileID,0,@CreateDate,@Finished,'','','','','','','','','','','') SET @Identity = SCOPE_IDENTITY()")


        End Try

        'ver1005
        Try
            If rs.State <> ADODB.ObjectStateEnum.adStateClosed Then rs.Close()
            rs.Open("Select * from MQ_Material2nd", glAdoInst.SelConnection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        Catch ex As Exception
            glAdoInst.SelConnection.Execute( _
            "CREATE TABLE [dbo].[MQ_Material2nd]([MaterialNo2nd] [nvarchar](150) NOT NULL," & _
            "[MaterialRes1] [nvarchar](150) NULL,[MaterialRes2] [nvarchar](150) NULL," & _
            "CONSTRAINT [PK_MQ_Material2nd] PRIMARY KEY CLUSTERED " & _
            "([MaterialNo2nd] Asc)WITH (PAD_INDEX  = OFF, IGNORE_DUP_KEY = OFF) ON [PRIMARY]) ON [PRIMARY]")
            glAdoInst.SelConnection.Execute("Insert Into MQ_Material2nd (MaterialNo2nd) values ('-A01')")
            glAdoInst.SelConnection.Execute("Insert Into MQ_Material2nd (MaterialNo2nd) values ('-A02')")
            glAdoInst.SelConnection.Execute("Insert Into MQ_Material2nd (MaterialNo2nd) values ('-B01')")


        End Try
        'ver1005 -1
        Try
            If rs.State <> ADODB.ObjectStateEnum.adStateClosed Then rs.Close()
            rs.Open("Select MaterialNo2nd from MQ_CheckIn Where MQCheckIn_ID = -1", glAdoInst.SelConnection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        Catch ex As Exception
            'alter MQ_CheckIn
            glAdoInst.SelConnection.Execute("ALTER TABLE dbo.MQ_CheckIn ADD MaterialNo2nd nvarchar(150) NULL")
            glAdoInst.SelConnection.Execute("Update dbo.MQ_CheckIn Set MaterialNo2nd = '' where MaterialNo2nd is null")
        End Try
        'ver1005 -2
        Try
            If rs.State <> ADODB.ObjectStateEnum.adStateClosed Then rs.Close()
            rs.Open("Select MaterialNo2nd from MQ_CheckOut Where CheckOutID = -1", glAdoInst.SelConnection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        Catch ex As Exception
            'alter MQ_CheckOut
            glAdoInst.SelConnection.Execute("ALTER TABLE dbo.MQ_CheckOut ADD MaterialNo2nd nvarchar(150) NULL")
            glAdoInst.SelConnection.Execute("Update dbo.MQ_CheckOut Set MaterialNo2nd = '' where MaterialNo2nd is null")
        End Try
        'ver1005 -3
        Try
            If rs.State <> ADODB.ObjectStateEnum.adStateClosed Then rs.Close()
            rs.Open("Select MaterialNo2nd from MQ_FileMaterial Where FMID = -1", glAdoInst.SelConnection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        Catch ex As Exception
            'alter MQ_FileMaterial
            glAdoInst.SelConnection.Execute("ALTER TABLE dbo.MQ_FileMaterial ADD MaterialNo2nd nvarchar(150) NULL")
            glAdoInst.SelConnection.Execute("Update dbo.MQ_FileMaterial Set MaterialNo2nd = '' where MaterialNo2nd is null")
        End Try
        'ver1005 -4
        Try
            If rs.State <> ADODB.ObjectStateEnum.adStateClosed Then rs.Close()
            rs.Open("Select MaterialNo2nd from MQ_FileName Where MQFileID = -1", glAdoInst.SelConnection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        Catch ex As Exception
            'alter MQ_FileName
            glAdoInst.SelConnection.Execute("ALTER TABLE dbo.MQ_FileName ADD MaterialNo2nd nvarchar(150) NULL")
            glAdoInst.SelConnection.Execute("Update dbo.MQ_FileName Set MaterialNo2nd = '' where MaterialNo2nd is null")
        End Try


        'ver1006   2012/10/12
        Try
            If rs.State <> ADODB.ObjectStateEnum.adStateClosed Then rs.Close()
            rs.Open("Select * from MQ_DBVer", glAdoInst.SelConnection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs.RecordCount > 0 Then
                rs.MoveFirst()
                If CInt(rs.Fields("DBVer").Value) < 1006 Then
                    glAdoInst.SelConnection.Execute( _
                           "ALTER PROCEDURE [dbo].[InsertMQ_Batch]  @MQFileID bigint,@Finished smallint,@CreateDate datetime,@MQBatchNo nvarchar(150)," & _
                           "@Level1 nvarchar(100),@Level2 nvarchar(100),@Level3 nvarchar(100),@Level4 nvarchar(100),@Level5 nvarchar(100),@Level6 nvarchar(100)," & _
                           "@Level7 nvarchar(100),@Level8 nvarchar(100),@Level9 nvarchar(100),@Level10 nvarchar(100), @Identity bigint OUT AS INSERT INTO MQ_Batch " & _
                           "(MQFileID,ExpectOutQty,CreateDate,Finished,MQBatchNo,Level1,Level2,Level3,Level4,Level5,Level6,Level7,Level8,Level9,Level10)" & _
                           " VALUES(@MQFileID,0,@CreateDate,@Finished,@MQBatchNo,@Level1,@Level2,@Level3,@Level4,@Level5,@Level6,@Level7,@Level8,@Level9,@Level10) " & _
                           "SET @Identity = SCOPE_IDENTITY()")
                    rs.Fields("DBVer").Value = "1006"
                    rs.Update()
                End If

            End If

        Catch ex As Exception
            MsgBox(Err.Description)
        End Try

        'ver1007   2012/11/12
        Try
            If rs.State <> ADODB.ObjectStateEnum.adStateClosed Then rs.Close()
            rs.Open("Select AnalysisResult from MQ_AutoNumber", glAdoInst.SelConnection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs.Close()
        Catch ex As Exception
            'alter MQ_AutoNumber, AnalysisResult
            glAdoInst.SelConnection.Execute("ALTER TABLE dbo.MQ_AutoNumber ADD AnalysisResult bigint null")
            glAdoInst.SelConnection.Execute("Update dbo.MQ_AutoNumber Set AnalysisResult = 1 where AnalysisResult is null")
            glAdoInst.SelConnection.Execute("Update dbo.MQ_DBVer Set DBVer = '1007'")
        End Try


        'ver1008   2012/11/24
        Try
            If rs.State <> ADODB.ObjectStateEnum.adStateClosed Then rs.Close()
            rs.Open("Select * from MQ_PassTable Where USERNAME = 'admin'", glAdoInst.SelConnection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs.Close()
        Catch ex As Exception
            'PassTable ----------------------
            cSQL = "CREATE TABLE [dbo].[MQ_PassTable](" & _
                    "[USERNAME] [nvarchar](20) NOT NULL," & _
                    "CONSTRAINT [PK_MQ_PassTables] PRIMARY KEY CLUSTERED " & _
                    "([USERNAME] Asc)WITH (PAD_INDEX  = OFF, IGNORE_DUP_KEY = OFF) ON [PRIMARY]) ON [PRIMARY]"
            glAdoInst.SelConnection.Execute(cSQL)

            For i = 1 To 100
                cSQL = "ALTER TABLE dbo.MQ_PassTable ADD MQRight" & i.ToString & " [nvarchar](7) null"
                glAdoInst.SelConnection.Execute(cSQL)
            Next
            '---------------------------------

            'MQ_RightTemplate ------------------
            cSQL = "CREATE TABLE [dbo].[MQ_RightTemplate](" & _
                    "[RID] [bigint] NOT NULL," & _
                    "[RName] [nvarchar](50) NULL," & _
                    "CONSTRAINT [PK_MQ_RightTemplate] PRIMARY KEY CLUSTERED " & _
                    "([RID] Asc)WITH (PAD_INDEX  = OFF, IGNORE_DUP_KEY = OFF) ON [PRIMARY]" & _
                    ") ON [PRIMARY]"

            glAdoInst.SelConnection.Execute(cSQL)

            For i = 1 To 100
                cSQL = "ALTER TABLE dbo.MQ_RightTemplate ADD MQRight" & i.ToString & " [nvarchar](7) null"
                glAdoInst.SelConnection.Execute(cSQL)
            Next
            '-------------------------------------

            '新增一筆 admin ---------------------
            If rs.State <> ADODB.ObjectStateEnum.adStateClosed Then rs.Close()
            cSQL = "Select * from MQ_PassTable"
            rs.Open(cSQL, glAdoInst.SelConnection, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)
            rs.AddNew()
            rs.Fields("USERNAME").Value = "admin"
            For i = 1 To 100
                rs.Fields("MQRight" & i.ToString).Value = "NDEVP--"
            Next
            rs.Update()
            '--------------------------------------

            glAdoInst.SelConnection.Execute("Update dbo.MQ_DBVer Set DBVer = '1008'")
        End Try

        'ver1008-1   2013/01/09
        '再次檢查是否有預存程序，有可能因為資料庫備份的關係，預存程序會消失。
        Try
            cSQL = "select * from sys.objects where object_id = OBJECT_ID(N'[dbo].InsertMQ_Batch')"
            If rs.State <> ADODB.ObjectStateEnum.adStateClosed Then rs.Close()
            rs.Open(cSQL, glAdoInst.SelConnection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs.RecordCount = 0 Then
                cSQL = "CREATE PROCEDURE [dbo].[InsertMQ_Batch]  @MQFileID bigint,@Finished smallint,@CreateDate datetime,@MQBatchNo nvarchar(150)," & _
                    "@Level1 nvarchar(100),@Level2 nvarchar(100),@Level3 nvarchar(100),@Level4 nvarchar(100),@Level5 nvarchar(100),@Level6 nvarchar(100)," & _
                    "@Level7 nvarchar(100),@Level8 nvarchar(100),@Level9 nvarchar(100),@Level10 nvarchar(100), @Identity bigint OUT AS INSERT INTO MQ_Batch " & _
                    "(MQFileID,ExpectOutQty,CreateDate,Finished,MQBatchNo,Level1,Level2,Level3,Level4,Level5,Level6,Level7,Level8,Level9,Level10) " & _
                    "VALUES(@MQFileID,0,@CreateDate,@Finished,@MQBatchNo,@Level1,@Level2,@Level3,@Level4,@Level5,@Level6,@Level7,@Level8,@Level9,@Level10) SET @Identity = SCOPE_IDENTITY()"
                glAdoInst.SelConnection.Execute(cSQL)
                glAdoInst.SelConnection.Execute("Update dbo.MQ_DBVer Set DBVer = '1008-1'")
            End If

        Catch ex As Exception

        End Try

        'ver1009   2013/01/25
        Try
            If rs.State <> ADODB.ObjectStateEnum.adStateClosed Then rs.Close()
            rs.Open("Select Alternative from MQ_CheckOut", glAdoInst.SelConnection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs.Close()
        Catch ex As Exception
            'alter MQ_AutoNumber, AnalysisResult
            glAdoInst.SelConnection.Execute("ALTER TABLE dbo.MQ_CheckOut ADD Alternative smallint null,OrigMaterialNo nvarchar(150) NULL,OrigMaterialNo2nd nvarchar(150) NULL")
            glAdoInst.SelConnection.Execute("Update dbo.MQ_CheckOut Set Alternative = 0, OrigMaterialNo= '', OrigMaterialNo2nd= '' where Alternative is null")
            glAdoInst.SelConnection.Execute("Update dbo.MQ_DBVer Set DBVer = '1009'")
        End Try

        'ver1010   2013/06/03
        Try
            If rs.State <> ADODB.ObjectStateEnum.adStateClosed Then rs.Close()
            rs.Open("Select * from MQ_NonView where NVID=0", glAdoInst.SelConnection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs.Close()
        Catch ex As Exception
            '建立 MQ_NonView table
            glAdoInst.SelConnection.Execute("CREATE TABLE [dbo].[MQ_NonView]([NID] [bigint] IDENTITY(1,1) NOT NULL,	[NVID] [bigint] NOT NULL,	[FLAG] [nchar](2) NOT NULL," & _
                                            "	[DATA] [nvarchar](50) NOT NULL, CONSTRAINT [PK_MQ_NonView] PRIMARY KEY CLUSTERED ([NID] ASC) WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF," & _
                                            " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) ON [PRIMARY]")
            '建立 MQ_FileTemplete
            glAdoInst.SelConnection.Execute("CREATE TABLE [dbo].[MQ_FileTemplate](	[FID] [bigint] IDENTITY(1,1) NOT NULL,	[FName] [nvarchar](50) NULL," & _
                                            "	[NVID] [bigint] NOT NULL, CONSTRAINT [PK_MQ_FileTemplate] PRIMARY KEY CLUSTERED (	[FID] ASC)WITH (PAD_INDEX  = OFF," & _
                                            " STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) ON [PRIMARY]")
            'Alter Table MQ_Passtable, NVID, FID
            glAdoInst.SelConnection.Execute("ALTER TABLE dbo.MQ_PassTable ADD NVID bigint null, FID bigint null")

            'alter MQ_AutoNumber, NVID
            glAdoInst.SelConnection.Execute("ALTER TABLE dbo.MQ_AutoNumber ADD NVID bigint null")
            glAdoInst.SelConnection.Execute("Update dbo.MQ_AutoNumber Set NVID = 1 where NVID is null")
            glAdoInst.SelConnection.Execute("Update dbo.MQ_DBVer Set DBVer = '1010'")

        End Try

        'ver1011   2013/06/11
        Try
            If rs.State <> ADODB.ObjectStateEnum.adStateClosed Then rs.Close()
            rs.Open("Select IsDefect from MQ_QCLEVEL Where QCLEVELID=-1", glAdoInst.SelConnection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs.Close()
        Catch ex As Exception

            'alter MQ_AutoNumber, NVID
            glAdoInst.SelConnection.Execute("ALTER TABLE dbo.MQ_QCLEVEL ADD IsDefect int null")
            glAdoInst.SelConnection.Execute("Update dbo.MQ_QCLEVEL Set IsDefect = 2 where IsDefect IS NULL")
            glAdoInst.SelConnection.Execute("Update dbo.MQ_QCLEVEL Set IsDefect = 1 where (QCLevelName ='不合格' or QCLevelName ='不良' or QCLevelName ='報廢')")
            glAdoInst.SelConnection.Execute("Update dbo.MQ_QCLEVEL Set IsDefect = 0 where (QCLevelName ='內控合格' or QCLevelName ='合格' or QCLevelName ='特採')")

            glAdoInst.SelConnection.Execute("Update dbo.MQ_DBVer Set DBVer = '1011'")

        End Try

        'ver1012
        Try
            If rs.State <> ADODB.ObjectStateEnum.adStateClosed Then rs.Close()
            rs.Open("Select DeftName from MQ_CheckIn Where MQCheckIn_ID = -1", glAdoInst.SelConnection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        Catch ex As Exception
            'alter MQ_CheckIn
            glAdoInst.SelConnection.Execute("ALTER TABLE dbo.MQ_CheckIn ADD DeftName nvarchar(50) NULL")
            glAdoInst.SelConnection.Execute("ALTER TABLE dbo.MQ_CheckIn ADD DeftID nvarchar(50) NULL")
            glAdoInst.SelConnection.Execute("Update dbo.MQ_CheckIn Set DeftName = '', DeftID = '' where DeftName is null")
            glAdoInst.SelConnection.Execute("ALTER TABLE dbo.MQ_CheckOut ADD DeftName nvarchar(50) NULL")
            glAdoInst.SelConnection.Execute("ALTER TABLE dbo.MQ_CheckOut ADD DeftID nvarchar(50) NULL")
            glAdoInst.SelConnection.Execute("Update dbo.MQ_CheckOut Set DeftName = '', DeftID = '' where DeftName is null")
            glAdoInst.SelConnection.Execute("Update dbo.MQ_DBVer Set DBVer = '1012'")
        End Try

        'ver1012 資料庫格式未變更
        Try
            If rs.State <> ADODB.ObjectStateEnum.adStateClosed Then rs.Close()
            rs.Open("Select MQFileName from MQ_FileName where MQResFile1 is NULL", glAdoInst.SelConnection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs.RecordCount > 0 Then
                glAdoInst.SelConnection.Execute("Update dbo.MQ_FileName Set MQResFile1 = 'False' where MQResFile1 is null")
            End If

        Catch ex As Exception


        End Try


        'ver1013
        Try
            If rs.State <> ADODB.ObjectStateEnum.adStateClosed Then rs.Close()
            rs.Open("Select UserDefine1 from MQ_CheckOut Where CheckOutID = -1", glAdoInst.SelConnection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        Catch ex As Exception
            'alter MQ_CheckIn
            glAdoInst.SelConnection.Execute("ALTER TABLE dbo.MQ_CheckOut ADD UserDefine1 nvarchar(100) NULL,UserDefine2 nvarchar(100) NULL,UserDefine3 nvarchar(100) NULL,UserDefine4 nvarchar(100) NULL,UserDefine5 nvarchar(100) NULL")
            glAdoInst.SelConnection.Execute("Delete From MQ_UserDefine Where UserType = 'CHECKOUT'")
            glAdoInst.SelConnection.Execute("Insert into MQ_UserDefine (UserType, Uno, IsTopLevel, Enabled, UserData) values ('CHECKOUT',1, -1,0,'發料單號')")
            glAdoInst.SelConnection.Execute("Insert into MQ_UserDefine (UserType, Uno, IsTopLevel, Enabled, UserData) values ('CHECKOUT',2, -1,0,'檢驗批號')")
            glAdoInst.SelConnection.Execute("Insert into MQ_UserDefine (UserType, Uno, IsTopLevel, Enabled, UserData) values ('CHECKOUT',3, -1,0,'出入庫批號')")
            glAdoInst.SelConnection.Execute("Insert into MQ_UserDefine (UserType, Uno, IsTopLevel, Enabled, UserData) values ('CHECKOUT',4, -1,0,'')")
            glAdoInst.SelConnection.Execute("Insert into MQ_UserDefine (UserType, Uno, IsTopLevel, Enabled, UserData) values ('CHECKOUT',5, -1,0,'')")
            glAdoInst.SelConnection.Execute("Update dbo.MQ_DBVer Set DBVer = '1013'")
        End Try


        'ver1014   2013/11/27
        Try
            If rs.State <> ADODB.ObjectStateEnum.adStateClosed Then rs.Close()
            rs.Open("Select BarCode from MQ_LEVELNAME Where ACTIVED=1", glAdoInst.SelConnection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs.Close()
        Catch ex As Exception

            'alter MQ_AutoNumber, NVID
            glAdoInst.SelConnection.Execute("ALTER TABLE dbo.MQ_LEVELNAME ADD BarCode int null")
            glAdoInst.SelConnection.Execute("Update dbo.MQ_LEVELNAME Set BarCode = 0 where BarCode IS NULL")

            glAdoInst.SelConnection.Execute("Update dbo.MQ_DBVer Set DBVer = '1014'")

        End Try


        'ver1015   2014/1/14
        Try
            If rs.State <> ADODB.ObjectStateEnum.adStateClosed Then rs.Close()
            rs.Open("Select * from MQ_Recorder Where RecID= -1", glAdoInst.SelConnection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs.Close()
        Catch ex As Exception

            'alter MQ_AutoNumber, NVID
            glAdoInst.SelConnection.Execute("CREATE TABLE [MQ_Recorder]([RecID] [bigint] IDENTITY(1,1) NOT NULL,[Type] [nvarchar](50) NULL,	[MQFileID] [bigint] NULL," & _
                                            "[MQBatchID] [bigint] NULL,	[MQCheckIn_ID] [bigint] NULL,[CheckOutID] [bigint] NULL,[EXMQBatchID] [bigint] NULL,[ID1] [bigint] NULL," & _
                                            "[ID2] [bigint] NULL,[ID3] [bigint] NULL,[DESC1] [nvarchar](50) NULL,[DESC2] [nvarchar](50) NULL,[DESC3] [nvarchar](50) NULL," & _
                                            "[DESC4] [nvarchar](50) NULL,[DESC5] [nvarchar](50) NULL,[PC] [nvarchar](50) NULL,[Time] [datetime] NULL, " & _
                                            "CONSTRAINT [PK_MQ_Recorder] PRIMARY KEY CLUSTERED ([RecID] ASC)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF," & _
                                            " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) ON [PRIMARY]")


            glAdoInst.SelConnection.Execute("Update dbo.MQ_DBVer Set DBVer = '1015'")

        End Try

        'ver1016   2014/2/12
        Try
            If rs.State <> ADODB.ObjectStateEnum.adStateClosed Then rs.Close()
            rs.Open("Select * from MQ_CondiName Where CondiID= -1", glAdoInst.SelConnection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs.Close()
        Catch ex As Exception
            'Create Table MQ_CondiName
            glAdoInst.SelConnection.Execute("CREATE TABLE [dbo].[MQ_CondiName]([CondiID] [bigint] IDENTITY(1,1) NOT NULL," & _
                                            "[CondiName] [nvarchar](100) NULL,[ParentID] [bigint] NULL,[CondiLevel] [bigint] NULL," & _
                                            "[CondiPath] [nvarchar](150) NULL,[CondiType] [nvarchar](50) NULL," & _
                                            "CONSTRAINT [PK_MQ_CondiName] PRIMARY KEY CLUSTERED ([CondiID] ASC)WITH (PAD_INDEX  = OFF," & _
                                            "STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON)" & _
                                            "ON [PRIMARY]) ON [PRIMARY]")
            glAdoInst.SelConnection.Execute("Update dbo.MQ_DBVer Set DBVer = '1016'")
        End Try
        Try
            If rs.State <> ADODB.ObjectStateEnum.adStateClosed Then rs.Close()
            rs.Open("Select * from MQ_SearchCondition Where CondiID= -1", glAdoInst.SelConnection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs.Close()
        Catch ex As Exception
            'Create Table MQ_CondiName
            glAdoInst.SelConnection.Execute("CREATE TABLE [dbo].[MQ_SearchCondition]([SCID] [bigint] IDENTITY(1,1) NOT NULL," & _
                                            "[CondiID] [bigint] NOT NULL,[Type] [nvarchar](50) NOT NULL,[Value1] [nvarchar](100) NULL," & _
                                            "[Value2] [nvarchar](100) NULL,[Value3] [nvarchar](100) NULL,[Value4] [nvarchar](100) NULL," & _
                                            "[Value5] [nvarchar](100) NULL,CONSTRAINT [PK_MQ_SearchCondition] PRIMARY KEY CLUSTERED ([SCID] ASC)" & _
                                            "WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON," & _
                                            "ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) ON [PRIMARY]")
            glAdoInst.SelConnection.Execute("Update dbo.MQ_DBVer Set DBVer = '1016'")

        End Try

        'ver1017   2014/4/9
        '更改MQCheckOut  MQCheckIn_ID 由smallint 改為 bigint
        Try
            If rs.State <> ADODB.ObjectStateEnum.adStateClosed Then rs.Close()
            rs.Open("SELECT   dbo.sysobjects.name AS sTableName, dbo.syscolumns.name AS sColumnsName," & _
               "dbo.syscolumns.prec AS iColumnsLength," & _
                        "dbo.syscolumns.colorder AS iColumnsOrder," & _
                        "dbo.systypes.name + '' AS sColumnsType," & _
                        "dbo.syscolumns.isnullable AS iIsNull" & _
                        " FROM dbo.sysobjects INNER JOIN" & _
                        " dbo.syscolumns ON dbo.sysobjects.id = dbo.syscolumns.id INNER JOIN" & _
                        " dbo.systypes ON dbo.syscolumns.xusertype = dbo.systypes.xusertype" & _
                        " WHERE (dbo.sysobjects.xtype = 'U' and dbo.sysobjects.name = 'MQ_CheckOut' and  dbo.syscolumns.name = 'MQCheckIn_ID')", glAdoInst.SelConnection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs.RecordCount > 0 Then
                rs.MoveFirst()
                If rs.Fields("sColumnsType").Value.ToString = "smallint" Then
                    glAdoInst.SelConnection.Execute("ALTER TABLE dbo.MQ_Checkout  Alter Column MQCheckIn_ID bigint NULL;")
                    glAdoInst.SelConnection.Execute("Update dbo.MQ_DBVer Set DBVer = '1017'")
                End If
            End If
            rs.Close()
        Catch ex As Exception

            glAdoInst.SelConnection.Execute("Update dbo.MQ_DBVer Set DBVer = '1016'")

        End Try

        '----------------------2014/8/15
        '更新MQ_QCLEVEL 加入Inform欄位
        Try
            If rs.State <> ADODB.ObjectStateEnum.adStateClosed Then rs.Close()
            rs.Open("Select Inform from MQ_QCLEVEL Where QCLEVELID=-1", glAdoInst.SelConnection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs.Close()
        Catch ex As Exception

            'alter MQ_AutoNumber, NVID
            glAdoInst.SelConnection.Execute("ALTER TABLE dbo.MQ_QCLEVEL ADD Inform int null")
            glAdoInst.SelConnection.Execute("Update dbo.MQ_QCLEVEL Set Inform = 1 where Inform IS NULL")
            glAdoInst.SelConnection.Execute("Update dbo.MQ_QCLEVEL Set Inform = 0 where (QCLevelName ='不合格' or QCLevelName ='不良' or QCLevelName ='報廢')")
            glAdoInst.SelConnection.Execute("Update dbo.MQ_QCLEVEL Set Inform = 1 where (QCLevelName ='內控合格' or QCLevelName ='合格' or QCLevelName ='特採')")

            glAdoInst.SelConnection.Execute("Update dbo.MQ_DBVer Set DBVer = '1017'")

        End Try


        '更新MQ_QCLEVEL 加入Lock欄位
        Try
            If rs.State <> ADODB.ObjectStateEnum.adStateClosed Then rs.Close()
            rs.Open("Select Lock from MQ_QCLEVEL Where QCLEVELID=-1", glAdoInst.SelConnection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs.Close()
        Catch ex As Exception

            'alter MQ_AutoNumber, NVID
            glAdoInst.SelConnection.Execute("ALTER TABLE dbo.MQ_QCLEVEL ADD Lock int null")
            glAdoInst.SelConnection.Execute("Update dbo.MQ_QCLEVEL Set Lock = 1 where Lock IS NULL")
            glAdoInst.SelConnection.Execute("Update dbo.MQ_QCLEVEL Set Lock = 0 where (QCLevelName ='不合格' or QCLevelName ='不良' or QCLevelName ='報廢')")
            glAdoInst.SelConnection.Execute("Update dbo.MQ_QCLEVEL Set Lock = 1 where (QCLevelName ='內控合格' or QCLevelName ='合格' or QCLevelName ='特採')")

            glAdoInst.SelConnection.Execute("Update dbo.MQ_DBVer Set DBVer = '1018'")

        End Try


        '更新MQ_FileName 加入Memo欄位
        Try
            If rs.State <> ADODB.ObjectStateEnum.adStateClosed Then rs.Close()
            rs.Open("Select Memo from MQ_FileName Where MQFileID=-1", glAdoInst.SelConnection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs.Close()
        Catch ex As Exception
            glAdoInst.SelConnection.Execute("ALTER TABLE dbo.MQ_FileName ADD Memo nvarchar(255) null")

            glAdoInst.SelConnection.Execute("Update dbo.MQ_DBVer Set DBVer = '1019'")
        End Try

        '更新MQ_CheckIN 加入Lock欄位
        Try
            If rs.State <> ADODB.ObjectStateEnum.adStateClosed Then rs.Close()
            rs.Open("Select Lock from MQ_CheckIn where MQCheckIn_ID =-1", glAdoInst.SelConnection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs.Close()
        Catch ex As Exception
            glAdoInst.SelConnection.Execute("ALTER TABLE dbo.MQ_CheckIn ADD Lock nvarchar(150) null")

            glAdoInst.SelConnection.Execute("Update dbo.MQ_DBVer Set DBVer = '1020'")
        End Try


        '新增MQ_LockRecord 資料表
        Try
            If rs.State <> ADODB.ObjectStateEnum.adStateClosed Then rs.Close()
            rs.Open("Select * from MQ_LockRecord where MQCheckIn_ID=-1", glAdoInst.SelConnection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs.Close()
        Catch ex As Exception
            glAdoInst.SelConnection.Execute("CREATE TABLE MQ_LockRecord (MQCheckIn_ID bigint, CheckInDate datetime,USERNAME nvarchar(20),UnLockDate datetime)")
            glAdoInst.SelConnection.Execute("Update dbo.MQ_DBVer Set DBVer = '1021'")
        End Try
    End Sub

End Class
