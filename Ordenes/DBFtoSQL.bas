Attribute VB_Name = "SQLtoDBF"
'****************************************************************
'Microsoft SQL Server 2000
'Visual Basic file generated for DTS Package
'File Name: C:\Documents and Settings\administrador.LUNA\Escritorio\Nuevo paquete.bas
'Package Name: ExportarDbf
'Package Description: Descripci�n del paquete DTS
'Generated Date: 26/12/2001
'Generated Time: 11:25:41 a.m.
'****************************************************************

Option Explicit
'Public goPackageOld As New DTS.Package
'Public goPackage As DTS.Package2
'
'Sub EliminarArchivos()
'On Error Resume Next
'Dim Arbol As New FileSystemObject
'Dim Carpeta As Object
''Set Carpeta = Arbol.GetFolder("\\sistemas01\confecc2")
'Set Carpeta = Arbol.GetFolder("\\server01\confecc\sql")
''If Not Arbol.FileExists(Carpeta) Then
'    Arbol.DeleteFile Carpeta & "\*.dbf", True
''End If
'
'Exit Sub
'hand:
'ErrorHandler Err, "EliminarArchivos"
'End Sub
'
'Public Sub EjecutaMigracionSQLtoDBF()
''\\server01\confecc\sql
'On Error Resume Next
'    EliminarArchivos
'        Set goPackage = goPackageOld
'
'
'        goPackage.Name = "ExportarDbf"
'        goPackage.Description = "Descripci�n del paquete DTS"
'        goPackage.WriteCompletionStatusToNTEventLog = False
'        goPackage.FailOnError = False
'        goPackage.PackagePriorityClass = 2
'        goPackage.MaxConcurrentSteps = 4
'        goPackage.LineageOptions = 0
'        goPackage.UseTransaction = True
'        goPackage.TransactionIsolationLevel = 4096
'        goPackage.AutoCommitTransaction = True
'        goPackage.RepositoryMetadataOptions = 0
'        goPackage.UseOLEDBServiceComponents = True
'        goPackage.LogToSQLServer = False
'        goPackage.LogServerFlags = 0
'        goPackage.FailPackageOnLogFailure = False
'        goPackage.ExplicitGlobalVariables = False
'        goPackage.PackageType = 0
'
'
'Dim oConnProperty As DTS.OleDBProperty
'
''---------------------------------------------------------------------------
'' create package connection information
''---------------------------------------------------------------------------
'
'Dim oConnection As DTS.Connection2
'
''------------- a new connection defined below.
''For security purposes, the password is never scripted
'
'Set oConnection = goPackage.Connections.New("SQLOLEDB")
'
'        oConnection.ConnectionProperties("Persist Security Info") = True
'        oConnection.ConnectionProperties("User ID") = "sa"
'        oConnection.ConnectionProperties("Initial Catalog") = "LIVES"
'        oConnection.ConnectionProperties("Data Source") = "server02"
'        oConnection.ConnectionProperties("Application Name") = "Asistente para importaci�n/exportaci�n con DTS"
'
'        oConnection.Name = "Conexi�n1"
'        oConnection.ID = 1
'        oConnection.Reusable = True
'        oConnection.ConnectImmediate = False
'        oConnection.DataSource = "server02"
'        oConnection.UserID = "sa"
'        oConnection.ConnectionTimeout = 60
'        oConnection.Catalog = "LIVES"
'        oConnection.UseTrustedConnection = False
'        oConnection.UseDSL = False
'
'        'If you have a password for this connection, please uncomment and add your password below.
'        'oConnection.Password = "<put the password here>"
'
'goPackage.Connections.Add oConnection
'Set oConnection = Nothing
'
''------------- a new connection defined below.
''For security purposes, the password is never scripted
'
'Set oConnection = goPackage.Connections.New("Microsoft.Jet.OLEDB.4.0")
'
'        'oConnection.ConnectionProperties("Data Source") = "\\server01\confecc\sql"
'        oConnection.ConnectionProperties("Data Source") = "\\server01\confecc\sql"
'        oConnection.ConnectionProperties("Extended Properties") = "dBase III"
'
'        oConnection.Name = "Conexi�n2"
'        oConnection.ID = 2
'        oConnection.Reusable = True
'        oConnection.ConnectImmediate = False
'        oConnection.DataSource = "\\server01\confecc\sql"
'        oConnection.ConnectionTimeout = 60
'        oConnection.UseTrustedConnection = False
'        oConnection.UseDSL = False
'
'        'If you have a password for this connection, please uncomment and add your password below.
'        'oConnection.Password = "<put the password here>"
'
'goPackage.Connections.Add oConnection
'Set oConnection = Nothing
'
''------------- a new connection defined below.
''For security purposes, the password is never scripted
'
'Set oConnection = goPackage.Connections.New("SQLOLEDB")
'
'        oConnection.ConnectionProperties("Persist Security Info") = True
'        oConnection.ConnectionProperties("User ID") = "sa"
'        oConnection.ConnectionProperties("Initial Catalog") = "LIVES"
'        oConnection.ConnectionProperties("Data Source") = "server02"
'        oConnection.ConnectionProperties("Application Name") = "Asistente para importaci�n/exportaci�n con DTS"
'
'        oConnection.Name = "Conexi�n3"
'        oConnection.ID = 3
'        oConnection.Reusable = True
'        oConnection.ConnectImmediate = False
'        oConnection.DataSource = "server02"
'        oConnection.UserID = "sa"
'        oConnection.ConnectionTimeout = 60
'        oConnection.Catalog = "LIVES"
'        oConnection.UseTrustedConnection = False
'        oConnection.UseDSL = False
'
'        'If you have a password for this connection, please uncomment and add your password below.
'        'oConnection.Password = "<put the password here>"
'
'goPackage.Connections.Add oConnection
'Set oConnection = Nothing
'
''------------- a new connection defined below.
''For security purposes, the password is never scripted
'
'Set oConnection = goPackage.Connections.New("Microsoft.Jet.OLEDB.4.0")
'
'        'oConnection.ConnectionProperties("Data Source") = "\\server01\confecc\sql"
'        oConnection.ConnectionProperties("Data Source") = "\\server01\confecc\sql"
'        oConnection.ConnectionProperties("Extended Properties") = "dBase III"
'
'        oConnection.Name = "Conexi�n4"
'        oConnection.ID = 4
'        oConnection.Reusable = True
'        oConnection.ConnectImmediate = False
'        'oConnection.DataSource = "\\server01\confecc\sql"
'        oConnection.DataSource = "\\server01\confecc\sql"
'
'        oConnection.ConnectionTimeout = 60
'        oConnection.UseTrustedConnection = False
'        oConnection.UseDSL = False
'
'        'If you have a password for this connection, please uncomment and add your password below.
'        'oConnection.Password = "<put the password here>"
'
'goPackage.Connections.Add oConnection
'Set oConnection = Nothing
'
''---------------------------------------------------------------------------
'' create package steps information
''---------------------------------------------------------------------------
'
'Dim oStep As DTS.Step2
'Dim oPrecConstraint As DTS.PrecedenceConstraint
'
''------------- a new step defined below
'
'Set oStep = goPackage.Steps.New
'
'        oStep.Name = "Crear tabla cf_clie Paso"
'        oStep.Description = "Crear tabla cf_clie Paso"
'        oStep.ExecutionStatus = 1
'        oStep.TaskName = "Crear tabla cf_clie Tarea"
'        oStep.CommitSuccess = False
'        oStep.RollbackFailure = False
'        oStep.ScriptLanguage = "VBScript"
'        oStep.AddGlobalVariables = True
'        oStep.RelativePriority = 3
'        oStep.CloseConnection = False
'        oStep.ExecuteInMainThread = False
'        oStep.IsPackageDSORowset = False
'        oStep.JoinTransactionIfPresent = False
'        oStep.DisableStep = False
'        oStep.FailPackageOnError = False
'
'goPackage.Steps.Add oStep
'Set oStep = Nothing
'
''------------- a new step defined below
'
'Set oStep = goPackage.Steps.New
'
'        oStep.Name = "Copy Data from cf_clie to cf_clie Paso"
'        oStep.Description = "Copy Data from cf_clie to cf_clie Paso"
'        oStep.ExecutionStatus = 1
'        oStep.TaskName = "Copy Data from cf_clie to cf_clie Tarea"
'        oStep.CommitSuccess = False
'        oStep.RollbackFailure = False
'        oStep.ScriptLanguage = "VBScript"
'        oStep.AddGlobalVariables = True
'        oStep.RelativePriority = 3
'        oStep.CloseConnection = False
'        oStep.ExecuteInMainThread = True
'        oStep.IsPackageDSORowset = False
'        oStep.JoinTransactionIfPresent = False
'        oStep.DisableStep = False
'        oStep.FailPackageOnError = False
'
'goPackage.Steps.Add oStep
'Set oStep = Nothing
'
''------------- a new step defined below
'
'Set oStep = goPackage.Steps.New
'
'        oStep.Name = "Crear tabla CF_DES Paso"
'        oStep.Description = "Crear tabla CF_DES Paso"
'        oStep.ExecutionStatus = 1
'        oStep.TaskName = "Crear tabla CF_DES Tarea"
'        oStep.CommitSuccess = False
'        oStep.RollbackFailure = False
'        oStep.ScriptLanguage = "VBScript"
'        oStep.AddGlobalVariables = True
'        oStep.RelativePriority = 3
'        oStep.CloseConnection = False
'        oStep.ExecuteInMainThread = False
'        oStep.IsPackageDSORowset = False
'        oStep.JoinTransactionIfPresent = False
'        oStep.DisableStep = False
'        oStep.FailPackageOnError = False
'
'goPackage.Steps.Add oStep
'Set oStep = Nothing
'
''------------- a new step defined below
'
'Set oStep = goPackage.Steps.New
'
'        oStep.Name = "Copy Data from CF_DES to CF_DES Paso"
'        oStep.Description = "Copy Data from CF_DES to CF_DES Paso"
'        oStep.ExecutionStatus = 1
'        oStep.TaskName = "Copy Data from CF_DES to CF_DES Tarea"
'        oStep.CommitSuccess = False
'        oStep.RollbackFailure = False
'        oStep.ScriptLanguage = "VBScript"
'        oStep.AddGlobalVariables = True
'        oStep.RelativePriority = 3
'        oStep.CloseConnection = False
'        oStep.ExecuteInMainThread = True
'        oStep.IsPackageDSORowset = False
'        oStep.JoinTransactionIfPresent = False
'        oStep.DisableStep = False
'        oStep.FailPackageOnError = False
'
'goPackage.Steps.Add oStep
'Set oStep = Nothing
'
''------------- a new step defined below
'
'Set oStep = goPackage.Steps.New
'
'        oStep.Name = "Crear tabla cf_pedd Paso"
'        oStep.Description = "Crear tabla cf_pedd Paso"
'        oStep.ExecutionStatus = 1
'        oStep.TaskName = "Crear tabla cf_pedd Tarea"
'        oStep.CommitSuccess = False
'        oStep.RollbackFailure = False
'        oStep.ScriptLanguage = "VBScript"
'        oStep.AddGlobalVariables = True
'        oStep.RelativePriority = 3
'        oStep.CloseConnection = False
'        oStep.ExecuteInMainThread = False
'        oStep.IsPackageDSORowset = False
'        oStep.JoinTransactionIfPresent = False
'        oStep.DisableStep = False
'        oStep.FailPackageOnError = False
'
'goPackage.Steps.Add oStep
'Set oStep = Nothing
'
''------------- a new step defined below
'
'Set oStep = goPackage.Steps.New
'
'        oStep.Name = "Copy Data from cf_pedd to cf_pedd Paso"
'        oStep.Description = "Copy Data from cf_pedd to cf_pedd Paso"
'        oStep.ExecutionStatus = 1
'        oStep.TaskName = "Copy Data from cf_pedd to cf_pedd Tarea"
'        oStep.CommitSuccess = False
'        oStep.RollbackFailure = False
'        oStep.ScriptLanguage = "VBScript"
'        oStep.AddGlobalVariables = True
'        oStep.RelativePriority = 3
'        oStep.CloseConnection = False
'        oStep.ExecuteInMainThread = True
'        oStep.IsPackageDSORowset = False
'        oStep.JoinTransactionIfPresent = False
'        oStep.DisableStep = False
'        oStep.FailPackageOnError = False
'
'goPackage.Steps.Add oStep
'Set oStep = Nothing
'
''------------- a new step defined below
'
'Set oStep = goPackage.Steps.New
'
'        oStep.Name = "Crear tabla cf_pedi Paso"
'        oStep.Description = "Crear tabla cf_pedi Paso"
'        oStep.ExecutionStatus = 1
'        oStep.TaskName = "Crear tabla cf_pedi Tarea"
'        oStep.CommitSuccess = False
'        oStep.RollbackFailure = False
'        oStep.ScriptLanguage = "VBScript"
'        oStep.AddGlobalVariables = True
'        oStep.RelativePriority = 3
'        oStep.CloseConnection = False
'        oStep.ExecuteInMainThread = False
'        oStep.IsPackageDSORowset = False
'        oStep.JoinTransactionIfPresent = False
'        oStep.DisableStep = False
'        oStep.FailPackageOnError = False
'
'goPackage.Steps.Add oStep
'Set oStep = Nothing
'
''------------- a new step defined below
'
'Set oStep = goPackage.Steps.New
'
'        oStep.Name = "Copy Data from cf_pedi to cf_pedi Paso"
'        oStep.Description = "Copy Data from cf_pedi to cf_pedi Paso"
'        oStep.ExecutionStatus = 1
'        oStep.TaskName = "Copy Data from cf_pedi to cf_pedi Tarea"
'        oStep.CommitSuccess = False
'        oStep.RollbackFailure = False
'        oStep.ScriptLanguage = "VBScript"
'        oStep.AddGlobalVariables = True
'        oStep.RelativePriority = 3
'        oStep.CloseConnection = False
'        oStep.ExecuteInMainThread = True
'        oStep.IsPackageDSORowset = False
'        oStep.JoinTransactionIfPresent = False
'        oStep.DisableStep = False
'        oStep.FailPackageOnError = False
'
'goPackage.Steps.Add oStep
'Set oStep = Nothing
'
''------------- a precedence constraint for steps defined below
'
'Set oStep = goPackage.Steps("Copy Data from cf_clie to cf_clie Paso")
'Set oPrecConstraint = oStep.PrecedenceConstraints.New("Crear tabla cf_clie Paso")
'        oPrecConstraint.StepName = "Crear tabla cf_clie Paso"
'        oPrecConstraint.PrecedenceBasis = 0
'        oPrecConstraint.value = 4
'
'oStep.PrecedenceConstraints.Add oPrecConstraint
'Set oPrecConstraint = Nothing
'
''------------- a precedence constraint for steps defined below
'
'Set oStep = goPackage.Steps("Copy Data from CF_DES to CF_DES Paso")
'Set oPrecConstraint = oStep.PrecedenceConstraints.New("Crear tabla CF_DES Paso")
'        oPrecConstraint.StepName = "Crear tabla CF_DES Paso"
'        oPrecConstraint.PrecedenceBasis = 0
'        oPrecConstraint.value = 4
'
'oStep.PrecedenceConstraints.Add oPrecConstraint
'Set oPrecConstraint = Nothing
'
''------------- a precedence constraint for steps defined below
'
'Set oStep = goPackage.Steps("Copy Data from cf_pedd to cf_pedd Paso")
'Set oPrecConstraint = oStep.PrecedenceConstraints.New("Crear tabla cf_pedd Paso")
'        oPrecConstraint.StepName = "Crear tabla cf_pedd Paso"
'        oPrecConstraint.PrecedenceBasis = 0
'        oPrecConstraint.value = 4
'
'oStep.PrecedenceConstraints.Add oPrecConstraint
'Set oPrecConstraint = Nothing
'
''------------- a precedence constraint for steps defined below
'
'Set oStep = goPackage.Steps("Copy Data from cf_pedi to cf_pedi Paso")
'Set oPrecConstraint = oStep.PrecedenceConstraints.New("Crear tabla cf_pedi Paso")
'        oPrecConstraint.StepName = "Crear tabla cf_pedi Paso"
'        oPrecConstraint.PrecedenceBasis = 0
'        oPrecConstraint.value = 4
'
'oStep.PrecedenceConstraints.Add oPrecConstraint
'Set oPrecConstraint = Nothing
'
''---------------------------------------------------------------------------
'' create package tasks information
''---------------------------------------------------------------------------
'
''------------- call Task_Sub1 for task Crear tabla cf_clie Tarea (Crear tabla cf_clie Tarea)
'Call Task_Sub1(goPackage)
'
''------------- call Task_Sub2 for task Copy Data from cf_clie to cf_clie Tarea (Copy Data from cf_clie to cf_clie Tarea)
'Call Task_Sub2(goPackage)
'
''------------- call Task_Sub3 for task Crear tabla CF_DES Tarea (Crear tabla CF_DES Tarea)
'Call Task_Sub3(goPackage)
'
''------------- call Task_Sub4 for task Copy Data from CF_DES to CF_DES Tarea (Copy Data from CF_DES to CF_DES Tarea)
'Call Task_Sub4(goPackage)
'
''------------- call Task_Sub5 for task Crear tabla cf_pedd Tarea (Crear tabla cf_pedd Tarea)
'Call Task_Sub5(goPackage)
'
''------------- call Task_Sub6 for task Copy Data from cf_pedd to cf_pedd Tarea (Copy Data from cf_pedd to cf_pedd Tarea)
'Call Task_Sub6(goPackage)
'
''------------- call Task_Sub7 for task Crear tabla cf_pedi Tarea (Crear tabla cf_pedi Tarea)
'Call Task_Sub7(goPackage)
'
''------------- call Task_Sub8 for task Copy Data from cf_pedi to cf_pedi Tarea (Copy Data from cf_pedi to cf_pedi Tarea)
'Call Task_Sub8(goPackage)
'
''---------------------------------------------------------------------------
'' Save or execute package
''---------------------------------------------------------------------------
'
''goPackage.SaveToSQLServer  "(local)", "sa", ""
'FrmMensaje.Titulo = "Ejecutando la copia de datos........."
'goPackage.Execute
''tracePackageError goPackage
'FrmMensaje.Titulo = "Proceso Terminado"
'FrmMensaje.Refresca
'goPackage.UnInitialize
''to save a package instead of executing it, comment out the executing package line above and uncomment the saving package line
'Set goPackage = Nothing
'
'Set goPackageOld = Nothing
'Unload FrmMensaje
'Exit Sub
'hand:
'ErrorHandler Err, "Main"
'Unload FrmMensaje
'Set goPackage = Nothing
'Set goPackageOld = Nothing
'
'End Sub
'
'
'Public Sub EjecutaMigracionSQLtoDBF2()
''\\sistemas01\Confecc2
'On Error Resume Next
'    EliminarArchivos
'        Set goPackage = goPackageOld
'
'
'        goPackage.Name = "ExportarDbf"
'        goPackage.Description = "Descripci�n del paquete DTS"
'        goPackage.WriteCompletionStatusToNTEventLog = False
'        goPackage.FailOnError = False
'        goPackage.PackagePriorityClass = 2
'        goPackage.MaxConcurrentSteps = 4
'        goPackage.LineageOptions = 0
'        goPackage.UseTransaction = True
'        goPackage.TransactionIsolationLevel = 4096
'        goPackage.AutoCommitTransaction = True
'        goPackage.RepositoryMetadataOptions = 0
'        goPackage.UseOLEDBServiceComponents = True
'        goPackage.LogToSQLServer = False
'        goPackage.LogServerFlags = 0
'        goPackage.FailPackageOnLogFailure = False
'        goPackage.ExplicitGlobalVariables = False
'        goPackage.PackageType = 0
'
'
'Dim oConnProperty As DTS.OleDBProperty
'
''---------------------------------------------------------------------------
'' create package connection information
''---------------------------------------------------------------------------
'
'Dim oConnection As DTS.Connection2
'
''------------- a new connection defined below.
''For security purposes, the password is never scripted
'
'Set oConnection = goPackage.Connections.New("SQLOLEDB")
'
'        oConnection.ConnectionProperties("Persist Security Info") = True
'        oConnection.ConnectionProperties("User ID") = "sa"
'        oConnection.ConnectionProperties("Initial Catalog") = "LIVES"
'        oConnection.ConnectionProperties("Data Source") = "SERVER02"
'        oConnection.ConnectionProperties("Application Name") = "Asistente para importaci�n/exportaci�n con DTS"
'
'        oConnection.Name = "Conexi�n1"
'        oConnection.ID = 1
'        oConnection.Reusable = True
'        oConnection.ConnectImmediate = False
'        oConnection.DataSource = "SERVER02"
'        oConnection.UserID = "sa"
'        oConnection.ConnectionTimeout = 60
'        oConnection.Catalog = "LIVES"
'        oConnection.UseTrustedConnection = False
'        oConnection.UseDSL = False
'
'        'If you have a password for this connection, please uncomment and add your password below.
'        'oConnection.Password = "<put the password here>"
'
'goPackage.Connections.Add oConnection
'Set oConnection = Nothing
'
''------------- a new connection defined below.
''For security purposes, the password is never scripted
'
'Set oConnection = goPackage.Connections.New("Microsoft.Jet.OLEDB.4.0")
'
'        'oConnection.ConnectionProperties("Data Source") = "\\sistemas01\Confecc2"
'        oConnection.ConnectionProperties("Data Source") = "\\server01\confecc\sql"
'        oConnection.ConnectionProperties("Extended Properties") = "dBase III"
'
'        oConnection.Name = "Conexi�n2"
'        oConnection.ID = 2
'        oConnection.Reusable = True
'        oConnection.ConnectImmediate = False
'        oConnection.DataSource = "\\server01\confecc\sql"
'        oConnection.ConnectionTimeout = 60
'        oConnection.UseTrustedConnection = False
'        oConnection.UseDSL = False
'
'        'If you have a password for this connection, please uncomment and add your password below.
'        'oConnection.Password = "<put the password here>"
'
'goPackage.Connections.Add oConnection
'Set oConnection = Nothing
'
''------------- a new connection defined below.
''For security purposes, the password is never scripted
'
'Set oConnection = goPackage.Connections.New("SQLOLEDB")
'
'        oConnection.ConnectionProperties("Persist Security Info") = True
'        oConnection.ConnectionProperties("User ID") = "sa"
'        oConnection.ConnectionProperties("Initial Catalog") = "LIVES"
'        oConnection.ConnectionProperties("Data Source") = "SERVER02"
'        oConnection.ConnectionProperties("Application Name") = "Asistente para importaci�n/exportaci�n con DTS"
'
'        oConnection.Name = "Conexi�n3"
'        oConnection.ID = 3
'        oConnection.Reusable = True
'        oConnection.ConnectImmediate = False
'        oConnection.DataSource = "SERVER02"
'        oConnection.UserID = "sa"
'        oConnection.ConnectionTimeout = 60
'        oConnection.Catalog = "LIVES"
'        oConnection.UseTrustedConnection = False
'        oConnection.UseDSL = False
'
'        'If you have a password for this connection, please uncomment and add your password below.
'        'oConnection.Password = "<put the password here>"
'
'goPackage.Connections.Add oConnection
'Set oConnection = Nothing
'
''------------- a new connection defined below.
''For security purposes, the password is never scripted
'
'Set oConnection = goPackage.Connections.New("Microsoft.Jet.OLEDB.4.0")
'
'        'oConnection.ConnectionProperties("Data Source") = "\\sistemas01\Confecc2"
'        oConnection.ConnectionProperties("Data Source") = "\\server01\confecc\sql"
'        oConnection.ConnectionProperties("Extended Properties") = "dBase III"
'
'        oConnection.Name = "Conexi�n4"
'        oConnection.ID = 4
'        oConnection.Reusable = True
'        oConnection.ConnectImmediate = False
'        'oConnection.DataSource = "\\sistemas01\Confecc2"
'        oConnection.DataSource = "\\server01\confecc\sql"
'
'        oConnection.ConnectionTimeout = 60
'        oConnection.UseTrustedConnection = False
'        oConnection.UseDSL = False
'
'        'If you have a password for this connection, please uncomment and add your password below.
'        'oConnection.Password = "<put the password here>"
'
'goPackage.Connections.Add oConnection
'Set oConnection = Nothing
'
''---------------------------------------------------------------------------
'' create package steps information
''---------------------------------------------------------------------------
'
'Dim oStep As DTS.Step2
'Dim oPrecConstraint As DTS.PrecedenceConstraint
'
''------------- a new step defined below
'
'Set oStep = goPackage.Steps.New
'
'        oStep.Name = "Crear tabla cf_clie Paso"
'        oStep.Description = "Crear tabla cf_clie Paso"
'        oStep.ExecutionStatus = 1
'        oStep.TaskName = "Crear tabla cf_clie Tarea"
'        oStep.CommitSuccess = False
'        oStep.RollbackFailure = False
'        oStep.ScriptLanguage = "VBScript"
'        oStep.AddGlobalVariables = True
'        oStep.RelativePriority = 3
'        oStep.CloseConnection = False
'        oStep.ExecuteInMainThread = False
'        oStep.IsPackageDSORowset = False
'        oStep.JoinTransactionIfPresent = False
'        oStep.DisableStep = False
'        oStep.FailPackageOnError = False
'
'goPackage.Steps.Add oStep
'Set oStep = Nothing
'
''------------- a new step defined below
'
'Set oStep = goPackage.Steps.New
'
'        oStep.Name = "Copy Data from cf_clie to cf_clie Paso"
'        oStep.Description = "Copy Data from cf_clie to cf_clie Paso"
'        oStep.ExecutionStatus = 1
'        oStep.TaskName = "Copy Data from cf_clie to cf_clie Tarea"
'        oStep.CommitSuccess = False
'        oStep.RollbackFailure = False
'        oStep.ScriptLanguage = "VBScript"
'        oStep.AddGlobalVariables = True
'        oStep.RelativePriority = 3
'        oStep.CloseConnection = False
'        oStep.ExecuteInMainThread = True
'        oStep.IsPackageDSORowset = False
'        oStep.JoinTransactionIfPresent = False
'        oStep.DisableStep = False
'        oStep.FailPackageOnError = False
'
'goPackage.Steps.Add oStep
'Set oStep = Nothing
'
''------------- a new step defined below
'
'Set oStep = goPackage.Steps.New
'
'        oStep.Name = "Crear tabla CF_DES Paso"
'        oStep.Description = "Crear tabla CF_DES Paso"
'        oStep.ExecutionStatus = 1
'        oStep.TaskName = "Crear tabla CF_DES Tarea"
'        oStep.CommitSuccess = False
'        oStep.RollbackFailure = False
'        oStep.ScriptLanguage = "VBScript"
'        oStep.AddGlobalVariables = True
'        oStep.RelativePriority = 3
'        oStep.CloseConnection = False
'        oStep.ExecuteInMainThread = False
'        oStep.IsPackageDSORowset = False
'        oStep.JoinTransactionIfPresent = False
'        oStep.DisableStep = False
'        oStep.FailPackageOnError = False
'
'goPackage.Steps.Add oStep
'Set oStep = Nothing
'
''------------- a new step defined below
'
'Set oStep = goPackage.Steps.New
'
'        oStep.Name = "Copy Data from CF_DES to CF_DES Paso"
'        oStep.Description = "Copy Data from CF_DES to CF_DES Paso"
'        oStep.ExecutionStatus = 1
'        oStep.TaskName = "Copy Data from CF_DES to CF_DES Tarea"
'        oStep.CommitSuccess = False
'        oStep.RollbackFailure = False
'        oStep.ScriptLanguage = "VBScript"
'        oStep.AddGlobalVariables = True
'        oStep.RelativePriority = 3
'        oStep.CloseConnection = False
'        oStep.ExecuteInMainThread = True
'        oStep.IsPackageDSORowset = False
'        oStep.JoinTransactionIfPresent = False
'        oStep.DisableStep = False
'        oStep.FailPackageOnError = False
'
'goPackage.Steps.Add oStep
'Set oStep = Nothing
'
''------------- a new step defined below
'
'Set oStep = goPackage.Steps.New
'
'        oStep.Name = "Crear tabla cf_pedd Paso"
'        oStep.Description = "Crear tabla cf_pedd Paso"
'        oStep.ExecutionStatus = 1
'        oStep.TaskName = "Crear tabla cf_pedd Tarea"
'        oStep.CommitSuccess = False
'        oStep.RollbackFailure = False
'        oStep.ScriptLanguage = "VBScript"
'        oStep.AddGlobalVariables = True
'        oStep.RelativePriority = 3
'        oStep.CloseConnection = False
'        oStep.ExecuteInMainThread = False
'        oStep.IsPackageDSORowset = False
'        oStep.JoinTransactionIfPresent = False
'        oStep.DisableStep = False
'        oStep.FailPackageOnError = False
'
'goPackage.Steps.Add oStep
'Set oStep = Nothing
'
''------------- a new step defined below
'
'Set oStep = goPackage.Steps.New
'
'        oStep.Name = "Copy Data from cf_pedd to cf_pedd Paso"
'        oStep.Description = "Copy Data from cf_pedd to cf_pedd Paso"
'        oStep.ExecutionStatus = 1
'        oStep.TaskName = "Copy Data from cf_pedd to cf_pedd Tarea"
'        oStep.CommitSuccess = False
'        oStep.RollbackFailure = False
'        oStep.ScriptLanguage = "VBScript"
'        oStep.AddGlobalVariables = True
'        oStep.RelativePriority = 3
'        oStep.CloseConnection = False
'        oStep.ExecuteInMainThread = True
'        oStep.IsPackageDSORowset = False
'        oStep.JoinTransactionIfPresent = False
'        oStep.DisableStep = False
'        oStep.FailPackageOnError = False
'
'goPackage.Steps.Add oStep
'Set oStep = Nothing
'
''------------- a new step defined below
'
'Set oStep = goPackage.Steps.New
'
'        oStep.Name = "Crear tabla cf_pedi Paso"
'        oStep.Description = "Crear tabla cf_pedi Paso"
'        oStep.ExecutionStatus = 1
'        oStep.TaskName = "Crear tabla cf_pedi Tarea"
'        oStep.CommitSuccess = False
'        oStep.RollbackFailure = False
'        oStep.ScriptLanguage = "VBScript"
'        oStep.AddGlobalVariables = True
'        oStep.RelativePriority = 3
'        oStep.CloseConnection = False
'        oStep.ExecuteInMainThread = False
'        oStep.IsPackageDSORowset = False
'        oStep.JoinTransactionIfPresent = False
'        oStep.DisableStep = False
'        oStep.FailPackageOnError = False
'
'goPackage.Steps.Add oStep
'Set oStep = Nothing
'
''------------- a new step defined below
'
'Set oStep = goPackage.Steps.New
'
'        oStep.Name = "Copy Data from cf_pedi to cf_pedi Paso"
'        oStep.Description = "Copy Data from cf_pedi to cf_pedi Paso"
'        oStep.ExecutionStatus = 1
'        oStep.TaskName = "Copy Data from cf_pedi to cf_pedi Tarea"
'        oStep.CommitSuccess = False
'        oStep.RollbackFailure = False
'        oStep.ScriptLanguage = "VBScript"
'        oStep.AddGlobalVariables = True
'        oStep.RelativePriority = 3
'        oStep.CloseConnection = False
'        oStep.ExecuteInMainThread = True
'        oStep.IsPackageDSORowset = False
'        oStep.JoinTransactionIfPresent = False
'        oStep.DisableStep = False
'        oStep.FailPackageOnError = False
'
'goPackage.Steps.Add oStep
'Set oStep = Nothing
'
''------------- a precedence constraint for steps defined below
'
'Set oStep = goPackage.Steps("Copy Data from cf_clie to cf_clie Paso")
'Set oPrecConstraint = oStep.PrecedenceConstraints.New("Crear tabla cf_clie Paso")
'        oPrecConstraint.StepName = "Crear tabla cf_clie Paso"
'        oPrecConstraint.PrecedenceBasis = 0
'        oPrecConstraint.value = 4
'
'oStep.PrecedenceConstraints.Add oPrecConstraint
'Set oPrecConstraint = Nothing
'
''------------- a precedence constraint for steps defined below
'
'Set oStep = goPackage.Steps("Copy Data from CF_DES to CF_DES Paso")
'Set oPrecConstraint = oStep.PrecedenceConstraints.New("Crear tabla CF_DES Paso")
'        oPrecConstraint.StepName = "Crear tabla CF_DES Paso"
'        oPrecConstraint.PrecedenceBasis = 0
'        oPrecConstraint.value = 4
'
'oStep.PrecedenceConstraints.Add oPrecConstraint
'Set oPrecConstraint = Nothing
'
''------------- a precedence constraint for steps defined below
'
'Set oStep = goPackage.Steps("Copy Data from cf_pedd to cf_pedd Paso")
'Set oPrecConstraint = oStep.PrecedenceConstraints.New("Crear tabla cf_pedd Paso")
'        oPrecConstraint.StepName = "Crear tabla cf_pedd Paso"
'        oPrecConstraint.PrecedenceBasis = 0
'        oPrecConstraint.value = 4
'
'oStep.PrecedenceConstraints.Add oPrecConstraint
'Set oPrecConstraint = Nothing
'
''------------- a precedence constraint for steps defined below
'
'Set oStep = goPackage.Steps("Copy Data from cf_pedi to cf_pedi Paso")
'Set oPrecConstraint = oStep.PrecedenceConstraints.New("Crear tabla cf_pedi Paso")
'        oPrecConstraint.StepName = "Crear tabla cf_pedi Paso"
'        oPrecConstraint.PrecedenceBasis = 0
'        oPrecConstraint.value = 4
'
'oStep.PrecedenceConstraints.Add oPrecConstraint
'Set oPrecConstraint = Nothing
'
''---------------------------------------------------------------------------
'' create package tasks information
''---------------------------------------------------------------------------
'
''------------- call Task_Sub1 for task Crear tabla cf_clie Tarea (Crear tabla cf_clie Tarea)
'Call Task_Sub1(goPackage)
'
''------------- call Task_Sub2 for task Copy Data from cf_clie to cf_clie Tarea (Copy Data from cf_clie to cf_clie Tarea)
'Call Task_Sub2(goPackage)
'
''------------- call Task_Sub3 for task Crear tabla CF_DES Tarea (Crear tabla CF_DES Tarea)
'Call Task_Sub3(goPackage)
'
''------------- call Task_Sub4 for task Copy Data from CF_DES to CF_DES Tarea (Copy Data from CF_DES to CF_DES Tarea)
'Call Task_Sub4(goPackage)
'
''------------- call Task_Sub5 for task Crear tabla cf_pedd Tarea (Crear tabla cf_pedd Tarea)
'Call Task_Sub5(goPackage)
'
''------------- call Task_Sub6 for task Copy Data from cf_pedd to cf_pedd Tarea (Copy Data from cf_pedd to cf_pedd Tarea)
'Call Task_Sub6(goPackage)
'
''------------- call Task_Sub7 for task Crear tabla cf_pedi Tarea (Crear tabla cf_pedi Tarea)
'Call Task_Sub7(goPackage)
'
''------------- call Task_Sub8 for task Copy Data from cf_pedi to cf_pedi Tarea (Copy Data from cf_pedi to cf_pedi Tarea)
'Call Task_Sub8(goPackage)
'
''---------------------------------------------------------------------------
'' Save or execute package
''---------------------------------------------------------------------------
'
''goPackage.SaveToSQLServer  "(local)", "sa", ""
'FrmMensaje.Titulo = "Ejecutando la copia de datos........."
'goPackage.Execute
''tracePackageError goPackage
''FrmMensaje.Titulo = "Proceso Terminado"
''FrmMensaje.Refresca
'goPackage.UnInitialize
''to save a package instead of executing it, comment out the executing package line above and uncomment the saving package line
'Set goPackage = Nothing
'
'Set goPackageOld = Nothing
''Unload FrmMensaje
'Exit Sub
'hand:
'ErrorHandler Err, "Main"
'Unload FrmMensaje
'Set goPackage = Nothing
'Set goPackageOld = Nothing
'
'End Sub
'
'
'
''-----------------------------------------------------------------------------
'' error reporting using step.GetExecutionErrorInfo after execution
''-----------------------------------------------------------------------------
'Public Sub tracePackageError(oPackage As DTS.Package)
'Dim ErrorCode As Long
'Dim ErrorSource As String
'Dim ErrorDescription As String
'Dim ErrorHelpFile As String
'Dim ErrorHelpContext As Long
'Dim ErrorIDofInterfaceWithError As String
'Dim i As Integer
'
'        For i = 1 To oPackage.Steps.Count
'                If oPackage.Steps(i).ExecutionResult = DTSStepExecResult_Failure Then
'                        oPackage.Steps(i).GetExecutionErrorInfo ErrorCode, ErrorSource, ErrorDescription, _
'                                        ErrorHelpFile, ErrorHelpContext, ErrorIDofInterfaceWithError
'                        MsgBox oPackage.Steps(i).Name & " failed" & vbCrLf & ErrorSource & vbCrLf & ErrorDescription, vbInformation
'                End If
'        Next i
'
'End Sub
'
''------------- define Task_Sub1 for task Crear tabla cf_clie Tarea (Crear tabla cf_clie Tarea)
'Public Sub Task_Sub1(ByVal goPackage As Object)
'
'Dim oTask As DTS.Task
'Dim oLookup As DTS.Lookup
'
'Dim oCustomTask1 As DTS.ExecuteSQLTask2
'Set oTask = goPackage.Tasks.New("DTSExecuteCommandSQLTask")
'oTask.Name = "Crear tabla cf_clie Tarea"
'Set oCustomTask1 = oTask.CustomTask
'
'        oCustomTask1.Name = "Crear tabla cf_clie Tarea"
'        oCustomTask1.Description = "Crear tabla cf_clie Tarea"
'        oCustomTask1.SQLStatement = "CREATE TABLE `cf_clie` (" & vbCrLf
'        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "`BCLICODCLI` VarChar (5) , " & vbCrLf
'        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "`BCLINOMCLI` VarChar (35) , " & vbCrLf
'        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "`BCLIABRCLI` VarChar (10) , " & vbCrLf
'        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "`BCLIZONCLI` VarChar (3) , " & vbCrLf
'        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "`BCLICLIPRI` VarChar (5) " & vbCrLf
'        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & ")"
'        oCustomTask1.ConnectionID = 2
'        oCustomTask1.CommandTimeout = 0
'        oCustomTask1.OutputAsRecordset = False
'
'goPackage.Tasks.Add oTask
'Set oCustomTask1 = Nothing
'Set oTask = Nothing
'
'End Sub
'
''------------- define Task_Sub2 for task Copy Data from cf_clie to cf_clie Tarea (Copy Data from cf_clie to cf_clie Tarea)
'Public Sub Task_Sub2(ByVal goPackage As Object)
'
'Dim oTask As DTS.Task
'Dim oLookup As DTS.Lookup
'
'Dim oCustomTask2 As DTS.DataPumpTask2
'Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
'oTask.Name = "Copy Data from cf_clie to cf_clie Tarea"
'Set oCustomTask2 = oTask.CustomTask
'
'        oCustomTask2.Name = "Copy Data from cf_clie to cf_clie Tarea"
'        oCustomTask2.Description = "Copy Data from cf_clie to cf_clie Tarea"
'        oCustomTask2.SourceConnectionID = 1
'        oCustomTask2.SourceSQLStatement = "select [BCLICODCLI],[BCLINOMCLI],[BCLIABRCLI],[BCLIZONCLI],[BCLICLIPRI] from [LIVES].[dbo].[cf_clie]"
'        oCustomTask2.DestinationConnectionID = 2
'        oCustomTask2.DestinationObjectName = "cf_clie"
'        oCustomTask2.ProgressRowCount = 1000
'        oCustomTask2.MaximumErrorCount = 0
'        oCustomTask2.FetchBufferSize = 1
'        oCustomTask2.UseFastLoad = True
'        oCustomTask2.InsertCommitSize = 0
'        oCustomTask2.ExceptionFileColumnDelimiter = "|"
'        oCustomTask2.ExceptionFileRowDelimiter = vbCrLf
'        oCustomTask2.AllowIdentityInserts = False
'        oCustomTask2.FirstRow = 0
'        oCustomTask2.LastRow = 0
'        oCustomTask2.FastLoadOptions = 2
'        oCustomTask2.ExceptionFileOptions = 1
'        oCustomTask2.DataPumpOptions = 0
'
'Call oCustomTask2_Trans_Sub1(oCustomTask2)
'
'
'goPackage.Tasks.Add oTask
'Set oCustomTask2 = Nothing
'Set oTask = Nothing
'
'End Sub
'
'Public Sub oCustomTask2_Trans_Sub1(ByVal oCustomTask2 As Object)
'
'        Dim oTransformation As DTS.Transformation2
'        Dim oTransProps As DTS.Properties
'        Dim oColumn As DTS.Column
'        Set oTransformation = oCustomTask2.Transformations.New("DTS.DataPumpTransformCopy")
'                oTransformation.Name = "DirectCopyXform"
'                oTransformation.TransformFlags = 63
'                oTransformation.ForceSourceBlobsBuffered = 0
'                oTransformation.ForceBlobsInMemory = False
'                oTransformation.InMemoryBlobSize = 1048576
'                oTransformation.TransformPhases = 4
'
'                Set oColumn = oTransformation.SourceColumns.New("BCLICODCLI", 1)
'                        oColumn.Name = "BCLICODCLI"
'                        oColumn.Ordinal = 1
'                        oColumn.Flags = 104
'                        oColumn.Size = 5
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BCLINOMCLI", 2)
'                        oColumn.Name = "BCLINOMCLI"
'                        oColumn.Ordinal = 2
'                        oColumn.Flags = 104
'                        oColumn.Size = 35
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BCLIABRCLI", 3)
'                        oColumn.Name = "BCLIABRCLI"
'                        oColumn.Ordinal = 3
'                        oColumn.Flags = 104
'                        oColumn.Size = 10
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BCLIZONCLI", 4)
'                        oColumn.Name = "BCLIZONCLI"
'                        oColumn.Ordinal = 4
'                        oColumn.Flags = 104
'                        oColumn.Size = 3
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BCLICLIPRI", 5)
'                        oColumn.Name = "BCLICLIPRI"
'                        oColumn.Ordinal = 5
'                        oColumn.Flags = 120
'                        oColumn.Size = 5
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BCLICODCLI", 1)
'                        oColumn.Name = "BCLICODCLI"
'                        oColumn.Ordinal = 1
'                        oColumn.Flags = 104
'                        oColumn.Size = 5
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BCLINOMCLI", 2)
'                        oColumn.Name = "BCLINOMCLI"
'                        oColumn.Ordinal = 2
'                        oColumn.Flags = 104
'                        oColumn.Size = 35
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BCLIABRCLI", 3)
'                        oColumn.Name = "BCLIABRCLI"
'                        oColumn.Ordinal = 3
'                        oColumn.Flags = 104
'                        oColumn.Size = 10
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BCLIZONCLI", 4)
'                        oColumn.Name = "BCLIZONCLI"
'                        oColumn.Ordinal = 4
'                        oColumn.Flags = 104
'                        oColumn.Size = 3
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BCLICLIPRI", 5)
'                        oColumn.Name = "BCLICLIPRI"
'                        oColumn.Ordinal = 5
'                        oColumn.Flags = 120
'                        oColumn.Size = 5
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'        Set oTransProps = oTransformation.TransformServerProperties
'
'
'        Set oTransProps = Nothing
'
'        oCustomTask2.Transformations.Add oTransformation
'        Set oTransformation = Nothing
'
'End Sub
'
''------------- define Task_Sub3 for task Crear tabla CF_DES Tarea (Crear tabla CF_DES Tarea)
'Public Sub Task_Sub3(ByVal goPackage As Object)
'
'Dim oTask As DTS.Task
'Dim oLookup As DTS.Lookup
'
'Dim oCustomTask3 As DTS.ExecuteSQLTask2
'Set oTask = goPackage.Tasks.New("DTSExecuteCommandSQLTask")
'oTask.Name = "Crear tabla CF_DES Tarea"
'Set oCustomTask3 = oTask.CustomTask
'
'        oCustomTask3.Name = "Crear tabla CF_DES Tarea"
'        oCustomTask3.Description = "Crear tabla CF_DES Tarea"
'        oCustomTask3.SQLStatement = "CREATE TABLE `CF_DES` (" & vbCrLf
'        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "`BDESCODDES` VarChar (3) , " & vbCrLf
'        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "`BDESDESDES` VarChar (15) , " & vbCrLf
'        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "`BDESABRDES` VarChar (3) " & vbCrLf
'        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & ")"
'        oCustomTask3.ConnectionID = 4
'        oCustomTask3.CommandTimeout = 0
'        oCustomTask3.OutputAsRecordset = False
'
'goPackage.Tasks.Add oTask
'Set oCustomTask3 = Nothing
'Set oTask = Nothing
'
'End Sub
'
''------------- define Task_Sub4 for task Copy Data from CF_DES to CF_DES Tarea (Copy Data from CF_DES to CF_DES Tarea)
'Public Sub Task_Sub4(ByVal goPackage As Object)
'
'Dim oTask As DTS.Task
'Dim oLookup As DTS.Lookup
'
'Dim oCustomTask4 As DTS.DataPumpTask2
'Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
'oTask.Name = "Copy Data from CF_DES to CF_DES Tarea"
'Set oCustomTask4 = oTask.CustomTask
'
'        oCustomTask4.Name = "Copy Data from CF_DES to CF_DES Tarea"
'        oCustomTask4.Description = "Copy Data from CF_DES to CF_DES Tarea"
'        oCustomTask4.SourceConnectionID = 3
'        oCustomTask4.SourceSQLStatement = "select [BDESCODDES],[BDESDESDES],[BDESABRDES] from [LIVES].[dbo].[CF_DES]"
'        oCustomTask4.DestinationConnectionID = 4
'        oCustomTask4.DestinationObjectName = "CF_DES"
'        oCustomTask4.ProgressRowCount = 1000
'        oCustomTask4.MaximumErrorCount = 0
'        oCustomTask4.FetchBufferSize = 1
'        oCustomTask4.UseFastLoad = True
'        oCustomTask4.InsertCommitSize = 0
'        oCustomTask4.ExceptionFileColumnDelimiter = "|"
'        oCustomTask4.ExceptionFileRowDelimiter = vbCrLf
'        oCustomTask4.AllowIdentityInserts = False
'        oCustomTask4.FirstRow = 0
'        oCustomTask4.LastRow = 0
'        oCustomTask4.FastLoadOptions = 2
'        oCustomTask4.ExceptionFileOptions = 1
'        oCustomTask4.DataPumpOptions = 0
'
'Call oCustomTask4_Trans_Sub1(oCustomTask4)
'
'
'goPackage.Tasks.Add oTask
'Set oCustomTask4 = Nothing
'Set oTask = Nothing
'
'End Sub
'
'Public Sub oCustomTask4_Trans_Sub1(ByVal oCustomTask4 As Object)
'
'        Dim oTransformation As DTS.Transformation2
'        Dim oTransProps As DTS.Properties
'        Dim oColumn As DTS.Column
'        Set oTransformation = oCustomTask4.Transformations.New("DTS.DataPumpTransformCopy")
'                oTransformation.Name = "DirectCopyXform"
'                oTransformation.TransformFlags = 63
'                oTransformation.ForceSourceBlobsBuffered = 0
'                oTransformation.ForceBlobsInMemory = False
'                oTransformation.InMemoryBlobSize = 1048576
'                oTransformation.TransformPhases = 4
'
'                Set oColumn = oTransformation.SourceColumns.New("BDESCODDES", 1)
'                        oColumn.Name = "BDESCODDES"
'                        oColumn.Ordinal = 1
'                        oColumn.Flags = 120
'                        oColumn.Size = 3
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BDESDESDES", 2)
'                        oColumn.Name = "BDESDESDES"
'                        oColumn.Ordinal = 2
'                        oColumn.Flags = 120
'                        oColumn.Size = 15
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BDESABRDES", 3)
'                        oColumn.Name = "BDESABRDES"
'                        oColumn.Ordinal = 3
'                        oColumn.Flags = 120
'                        oColumn.Size = 3
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BDESCODDES", 1)
'                        oColumn.Name = "BDESCODDES"
'                        oColumn.Ordinal = 1
'                        oColumn.Flags = 120
'                        oColumn.Size = 3
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BDESDESDES", 2)
'                        oColumn.Name = "BDESDESDES"
'                        oColumn.Ordinal = 2
'                        oColumn.Flags = 120
'                        oColumn.Size = 15
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BDESABRDES", 3)
'                        oColumn.Name = "BDESABRDES"
'                        oColumn.Ordinal = 3
'                        oColumn.Flags = 120
'                        oColumn.Size = 3
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'        Set oTransProps = oTransformation.TransformServerProperties
'
'
'        Set oTransProps = Nothing
'
'        oCustomTask4.Transformations.Add oTransformation
'        Set oTransformation = Nothing
'
'End Sub
'
''------------- define Task_Sub5 for task Crear tabla cf_pedd Tarea (Crear tabla cf_pedd Tarea)
'Public Sub Task_Sub5(ByVal goPackage As Object)
'
'Dim oTask As DTS.Task
'Dim oLookup As DTS.Lookup
'
'Dim oCustomTask5 As DTS.ExecuteSQLTask2
'Set oTask = goPackage.Tasks.New("DTSExecuteCommandSQLTask")
'oTask.Name = "Crear tabla cf_pedd Tarea"
'Set oCustomTask5 = oTask.CustomTask
'
'        oCustomTask5.Name = "Crear tabla cf_pedd Tarea"
'        oCustomTask5.Description = "Crear tabla cf_pedd Tarea"
'        oCustomTask5.SQLStatement = "CREATE TABLE `cf_pedd` (" & vbCrLf
'        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "`BPDITIPPED` VarChar (1) , " & vbCrLf
'        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "`BPDICODPED` VarChar (5) , " & vbCrLf
'        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "`BPDINUMPRE` VarChar (3) , " & vbCrLf
'        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "`BPDIDESPRE` VarChar (35) , " & vbCrLf
'        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "`BPDICANPRE` Long , " & vbCrLf
'        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "`BPDICODETI` VarChar (20) , " & vbCrLf
'        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "`BPDICANPRR` Long , " & vbCrLf
'        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "`BPDICODTAL` VarChar (4) , " & vbCrLf
'        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "`BPDINUMVER` VarChar (3) , " & vbCrLf
'        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "`BPDIREFCOL` VarChar (8) , " & vbCrLf
'        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "`BPDICANDES` Long " & vbCrLf
'        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & ")"
'        oCustomTask5.ConnectionID = 2
'        oCustomTask5.CommandTimeout = 0
'        oCustomTask5.OutputAsRecordset = False
'
'goPackage.Tasks.Add oTask
'Set oCustomTask5 = Nothing
'Set oTask = Nothing
'
'End Sub
'
''------------- define Task_Sub6 for task Copy Data from cf_pedd to cf_pedd Tarea (Copy Data from cf_pedd to cf_pedd Tarea)
'Public Sub Task_Sub6(ByVal goPackage As Object)
'
'Dim oTask As DTS.Task
'Dim oLookup As DTS.Lookup
'
'Dim oCustomTask6 As DTS.DataPumpTask2
'Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
'oTask.Name = "Copy Data from cf_pedd to cf_pedd Tarea"
'Set oCustomTask6 = oTask.CustomTask
'
'        oCustomTask6.Name = "Copy Data from cf_pedd to cf_pedd Tarea"
'        oCustomTask6.Description = "Copy Data from cf_pedd to cf_pedd Tarea"
'        oCustomTask6.SourceConnectionID = 1
'        oCustomTask6.SourceSQLStatement = "select [BPDITIPPED],[BPDICODPED],[BPDINUMPRE],[BPDIDESPRE],[BPDICANPRE],[BPDICODETI],[BPDICANPRR],[BPDICODTAL],[BPDINUMVER],[BPDIREFCOL],[BPDICANDES] from [LIVES].[dbo].[cf_pedd]"
'        oCustomTask6.DestinationConnectionID = 2
'        oCustomTask6.DestinationObjectName = "cf_pedd"
'        oCustomTask6.ProgressRowCount = 1000
'        oCustomTask6.MaximumErrorCount = 0
'        oCustomTask6.FetchBufferSize = 1
'        oCustomTask6.UseFastLoad = True
'        oCustomTask6.InsertCommitSize = 0
'        oCustomTask6.ExceptionFileColumnDelimiter = "|"
'        oCustomTask6.ExceptionFileRowDelimiter = vbCrLf
'        oCustomTask6.AllowIdentityInserts = False
'        oCustomTask6.FirstRow = 0
'        oCustomTask6.LastRow = 0
'        oCustomTask6.FastLoadOptions = 2
'        oCustomTask6.ExceptionFileOptions = 1
'        oCustomTask6.DataPumpOptions = 0
'
'Call oCustomTask6_Trans_Sub1(oCustomTask6)
'
'
'goPackage.Tasks.Add oTask
'Set oCustomTask6 = Nothing
'Set oTask = Nothing
'
'End Sub
'
'Public Sub oCustomTask6_Trans_Sub1(ByVal oCustomTask6 As Object)
'
'        Dim oTransformation As DTS.Transformation2
'        Dim oTransProps As DTS.Properties
'        Dim oColumn As DTS.Column
'        Set oTransformation = oCustomTask6.Transformations.New("DTS.DataPumpTransformCopy")
'                oTransformation.Name = "DirectCopyXform"
'                oTransformation.TransformFlags = 63
'                oTransformation.ForceSourceBlobsBuffered = 0
'                oTransformation.ForceBlobsInMemory = False
'                oTransformation.InMemoryBlobSize = 1048576
'                oTransformation.TransformPhases = 4
'
'                Set oColumn = oTransformation.SourceColumns.New("BPDITIPPED", 1)
'                        oColumn.Name = "BPDITIPPED"
'                        oColumn.Ordinal = 1
'                        oColumn.Flags = 120
'                        oColumn.Size = 1
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPDICODPED", 2)
'                        oColumn.Name = "BPDICODPED"
'                        oColumn.Ordinal = 2
'                        oColumn.Flags = 120
'                        oColumn.Size = 5
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPDINUMPRE", 3)
'                        oColumn.Name = "BPDINUMPRE"
'                        oColumn.Ordinal = 3
'                        oColumn.Flags = 120
'                        oColumn.Size = 3
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPDIDESPRE", 4)
'                        oColumn.Name = "BPDIDESPRE"
'                        oColumn.Ordinal = 4
'                        oColumn.Flags = 120
'                        oColumn.Size = 35
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPDICANPRE", 5)
'                        oColumn.Name = "BPDICANPRE"
'                        oColumn.Ordinal = 5
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 3
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPDICODETI", 6)
'                        oColumn.Name = "BPDICODETI"
'                        oColumn.Ordinal = 6
'                        oColumn.Flags = 120
'                        oColumn.Size = 10
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPDICANPRR", 7)
'                        oColumn.Name = "BPDICANPRR"
'                        oColumn.Ordinal = 7
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 3
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPDICODTAL", 8)
'                        oColumn.Name = "BPDICODTAL"
'                        oColumn.Ordinal = 8
'                        oColumn.Flags = 120
'                        oColumn.Size = 4
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPDINUMVER", 9)
'                        oColumn.Name = "BPDINUMVER"
'                        oColumn.Ordinal = 9
'                        oColumn.Flags = 120
'                        oColumn.Size = 3
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPDIREFCOL", 10)
'                        oColumn.Name = "BPDIREFCOL"
'                        oColumn.Ordinal = 10
'                        oColumn.Flags = 120
'                        oColumn.Size = 8
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPDICANDES", 11)
'                        oColumn.Name = "BPDICANDES"
'                        oColumn.Ordinal = 11
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 3
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPDITIPPED", 1)
'                        oColumn.Name = "BPDITIPPED"
'                        oColumn.Ordinal = 1
'                        oColumn.Flags = 120
'                        oColumn.Size = 1
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPDICODPED", 2)
'                        oColumn.Name = "BPDICODPED"
'                        oColumn.Ordinal = 2
'                        oColumn.Flags = 120
'                        oColumn.Size = 5
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPDINUMPRE", 3)
'                        oColumn.Name = "BPDINUMPRE"
'                        oColumn.Ordinal = 3
'                        oColumn.Flags = 120
'                        oColumn.Size = 3
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPDIDESPRE", 4)
'                        oColumn.Name = "BPDIDESPRE"
'                        oColumn.Ordinal = 4
'                        oColumn.Flags = 120
'                        oColumn.Size = 35
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPDICANPRE", 5)
'                        oColumn.Name = "BPDICANPRE"
'                        oColumn.Ordinal = 5
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 3
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPDICODETI", 6)
'                        oColumn.Name = "BPDICODETI"
'                        oColumn.Ordinal = 6
'                        oColumn.Flags = 120
'                        oColumn.Size = 10
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPDICANPRR", 7)
'                        oColumn.Name = "BPDICANPRR"
'                        oColumn.Ordinal = 7
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 3
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPDICODTAL", 8)
'                        oColumn.Name = "BPDICODTAL"
'                        oColumn.Ordinal = 8
'                        oColumn.Flags = 120
'                        oColumn.Size = 4
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPDINUMVER", 9)
'                        oColumn.Name = "BPDINUMVER"
'                        oColumn.Ordinal = 9
'                        oColumn.Flags = 120
'                        oColumn.Size = 3
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPDIREFCOL", 10)
'                        oColumn.Name = "BPDIREFCOL"
'                        oColumn.Ordinal = 10
'                        oColumn.Flags = 120
'                        oColumn.Size = 8
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPDICANDES", 11)
'                        oColumn.Name = "BPDICANDES"
'                        oColumn.Ordinal = 11
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 3
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'        Set oTransProps = oTransformation.TransformServerProperties
'
'
'        Set oTransProps = Nothing
'
'        oCustomTask6.Transformations.Add oTransformation
'        Set oTransformation = Nothing
'
'End Sub
'
''------------- define Task_Sub7 for task Crear tabla cf_pedi Tarea (Crear tabla cf_pedi Tarea)
'Public Sub Task_Sub7(ByVal goPackage As Object)
'
'Dim oTask As DTS.Task
'Dim oLookup As DTS.Lookup
'
'Dim oCustomTask7 As DTS.ExecuteSQLTask2
'Set oTask = goPackage.Tasks.New("DTSExecuteCommandSQLTask")
'oTask.Name = "Crear tabla cf_pedi Tarea"
'Set oCustomTask7 = oTask.CustomTask
'
'        oCustomTask7.Name = "Crear tabla cf_pedi Tarea"
'        oCustomTask7.Description = "Crear tabla cf_pedi Tarea"
'        oCustomTask7.SQLStatement = "CREATE TABLE `cf_pedi` (" & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1TIPPED` VarChar (1) , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1CODPED` VarChar (5) , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1CODCLI` VarChar (5) , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1CODEST` VarChar (9) , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1TIPTAR` VarChar (1) , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1CODTAR` VarChar (5) , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1CLATAR` VarChar (1) , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1CODPTA` VarChar (2) , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1ULTOPR` VarChar (3) , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1UNIPED` VarChar (3) , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1DESPED` VarChar (35) , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1INISAL` DateTime , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1FINSAL` DateTime , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1FECDES` DateTime , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1ESTPLA` VarChar (9) , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1ESTGEN` VarChar (9) , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1TIPAPI` VarChar (1) , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1TIPTEL` VarChar (1) , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1TIECOR` Long , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1FECNEG` DateTime , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1TIEBOR` Long , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1ESTCLI` VarChar (10) , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1PEDRET` VarChar (1) , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1PRIORI` VarChar (1) , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1FECOBJ` DateTime , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1DIVCLI` VarChar (3) , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1FECMOD` DateTime , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1TIPERE` VarChar (1) , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1CODERE` VarChar (6) , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1TIPIFI` VarChar (1) , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1CODACT` VarChar (2) , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1FECRET` DateTime , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1PURORD` VarChar (15) , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1CODDES` VarChar (3) , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1FECCAN` DateTime , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1FEULDE` DateTime , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1FECLIQ` DateTime , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1FECEMS` DateTime , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1MARCAS` VarChar (1) , " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "`BPE1CODVAR` VarChar (1) " & vbCrLf
'        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & ")"
'        oCustomTask7.ConnectionID = 4
'        oCustomTask7.CommandTimeout = 0
'        oCustomTask7.OutputAsRecordset = False
'
'goPackage.Tasks.Add oTask
'Set oCustomTask7 = Nothing
'Set oTask = Nothing
'
'End Sub
'
''------------- define Task_Sub8 for task Copy Data from cf_pedi to cf_pedi Tarea (Copy Data from cf_pedi to cf_pedi Tarea)
'Public Sub Task_Sub8(ByVal goPackage As Object)
'On Error Resume Next
'Dim oTask As DTS.Task
'Dim oLookup As DTS.Lookup
'
'Dim oCustomTask8 As DTS.DataPumpTask2
'Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
'oTask.Name = "Copy Data from cf_pedi to cf_pedi Tarea"
'Set oCustomTask8 = oTask.CustomTask
'
'        oCustomTask8.Name = "Copy Data from cf_pedi to cf_pedi Tarea"
'        oCustomTask8.Description = "Copy Data from cf_pedi to cf_pedi Tarea"
'        oCustomTask8.SourceConnectionID = 3
'        'oCustomTask8.SourceSQLStatement = "select [BPE1TIPPED],[BPE1CODPED],[BPE1CODCLI],[BPE1CODEST],[BPE1TIPTAR],[BPE1CODTAR],[BPE1CLATAR],[BPE1CODPTA],[BPE1ULTOPR],[BPE1UNIPED],[BPE1DESPED],[BPE1INISAL],[BPE1FINSAL],[BPE1FECDES],[BPE1ESTPLA],[BPE1ESTGEN],[BPE1TIPAPI],[BPE1TIPTEL],[BPE1TIECOR],["
'        'oCustomTask8.SourceSQLStatement = "select [BPE1TIPPED],[BPE1CODPED],[BPE1CODCLI],[BPE1CODEST],[BPE1TIPTAR],[BPE1CODTAR],[BPE1CLATAR],[BPE1CODPTA],[BPE1ULTOPR],[BPE1UNIPED],[BPE1DESPED],[BPE1INISAL],[BPE1FINSAL],[BPE1FECDES],[BPE1ESTPLA],[BPE1ESTGEN],[BPE1TIPAPI],[BPE1TIPTEL],[BPE1TIECOR],["
'        'oCustomTask8.SourceSQLStatement = oCustomTask8.SourceSQLStatement & "BPE1FECNEG],[BPE1TIEBOR],[BPE1ESTCLI],[BPE1PEDRET],[BPE1PRIORI],[BPE1FECOBJ],[BPE1DIVCLI],[BPE1FECMOD],[BPE1TIPERE],[BPE1CODERE],[BPE1TIPIFI],[BPE1CODACT],[BPE1FECRET],[BPE1PURORD],[BPE1CODDES],[BPE1FECCAN],[BPE1FEULDE],[BPE1FECLIQ],[BPE1FECEMS],[BPE1MARC"
'        oCustomTask8.SourceSQLStatement = "select [BPE1TIPPED],[BPE1CODPED],[BPE1CODCLI],[BPE1CODEST],[BPE1TIPTAR],[BPE1CODTAR],[BPE1CLATAR],[BPE1CODPTA],[BPE1ULTOPR],[BPE1UNIPED],[BPE1DESPED],[BPE1INISAL],[BPE1FINSAL],[BPE1FECDES],[BPE1ESTPLA],[BPE1ESTGEN],[BPE1TIPAPI],[BPE1TIPTEL],[BPE1TIECOR],["
'        oCustomTask8.SourceSQLStatement = oCustomTask8.SourceSQLStatement & "BPE1FECNEG],[BPE1TIEBOR],[BPE1ESTCLI],[BPE1PEDRET],[BPE1PRIORI],[BPE1FECOBJ],[BPE1DIVCLI],[BPE1FECMOD],[BPE1TIPERE],[BPE1CODERE],[BPE1TIPIFI],[BPE1CODACT],[BPE1FECRET],[BPE1PURORD],[BPE1CODDES],[BPE1FECCAN],[BPE1FEULDE],[BPE1FECLIQ],[BPE1FECEMS],[BPE1MARC"
'        oCustomTask8.SourceSQLStatement = oCustomTask8.SourceSQLStatement & "AS],[BPE1CODVAR] from [LIVES].[dbo].[cf_pedi]"
'
'        oCustomTask8.DestinationConnectionID = 4
'        oCustomTask8.DestinationObjectName = "cf_pedi"
'        oCustomTask8.ProgressRowCount = 1000
'        oCustomTask8.MaximumErrorCount = 0
'        oCustomTask8.FetchBufferSize = 1
'        oCustomTask8.UseFastLoad = True
'        oCustomTask8.InsertCommitSize = 0
'        oCustomTask8.ExceptionFileColumnDelimiter = "|"
'        oCustomTask8.ExceptionFileRowDelimiter = vbCrLf
'        oCustomTask8.AllowIdentityInserts = False
'        oCustomTask8.FirstRow = 0
'        oCustomTask8.LastRow = 0
'        oCustomTask8.FastLoadOptions = 2
'        oCustomTask8.ExceptionFileOptions = 1
'        oCustomTask8.DataPumpOptions = 0
'
'Call oCustomTask8_Trans_Sub1(oCustomTask8)
'
'
'goPackage.Tasks.Add oTask
'Set oCustomTask8 = Nothing
'Set oTask = Nothing
'
'End Sub
'
'Public Sub oCustomTask8_Trans_Sub1(ByVal oCustomTask8 As Object)
'
'        Dim oTransformation As DTS.Transformation2
'        Dim oTransProps As DTS.Properties
'        Dim oColumn As DTS.Column
'        Set oTransformation = oCustomTask8.Transformations.New("DTS.DataPumpTransformCopy")
'                oTransformation.Name = "DirectCopyXform"
'                oTransformation.TransformFlags = 63
'                oTransformation.ForceSourceBlobsBuffered = 0
'                oTransformation.ForceBlobsInMemory = False
'                oTransformation.InMemoryBlobSize = 1048576
'                oTransformation.TransformPhases = 4
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1TIPPED", 1)
'                        oColumn.Name = "BPE1TIPPED"
'                        oColumn.Ordinal = 1
'                        oColumn.Flags = 120
'                        oColumn.Size = 1
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1CODPED", 2)
'                        oColumn.Name = "BPE1CODPED"
'                        oColumn.Ordinal = 2
'                        oColumn.Flags = 120
'                        oColumn.Size = 5
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1CODCLI", 3)
'                        oColumn.Name = "BPE1CODCLI"
'                        oColumn.Ordinal = 3
'                        oColumn.Flags = 120
'                        oColumn.Size = 5
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1CODEST", 4)
'                        oColumn.Name = "BPE1CODEST"
'                        oColumn.Ordinal = 4
'                        oColumn.Flags = 120
'                        oColumn.Size = 9
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1TIPTAR", 5)
'                        oColumn.Name = "BPE1TIPTAR"
'                        oColumn.Ordinal = 5
'                        oColumn.Flags = 120
'                        oColumn.Size = 1
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1CODTAR", 6)
'                        oColumn.Name = "BPE1CODTAR"
'                        oColumn.Ordinal = 6
'                        oColumn.Flags = 120
'                        oColumn.Size = 5
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1CLATAR", 7)
'                        oColumn.Name = "BPE1CLATAR"
'                        oColumn.Ordinal = 7
'                        oColumn.Flags = 120
'                        oColumn.Size = 1
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1CODPTA", 8)
'                        oColumn.Name = "BPE1CODPTA"
'                        oColumn.Ordinal = 8
'                        oColumn.Flags = 120
'                        oColumn.Size = 2
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1ULTOPR", 9)
'                        oColumn.Name = "BPE1ULTOPR"
'                        oColumn.Ordinal = 9
'                        oColumn.Flags = 120
'                        oColumn.Size = 3
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1UNIPED", 10)
'                        oColumn.Name = "BPE1UNIPED"
'                        oColumn.Ordinal = 10
'                        oColumn.Flags = 120
'                        oColumn.Size = 3
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1DESPED", 11)
'                        oColumn.Name = "BPE1DESPED"
'                        oColumn.Ordinal = 11
'                        oColumn.Flags = 120
'                        oColumn.Size = 35
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1INISAL", 12)
'                        oColumn.Name = "BPE1INISAL"
'                        oColumn.Ordinal = 12
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 135
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1FINSAL", 13)
'                        oColumn.Name = "BPE1FINSAL"
'                        oColumn.Ordinal = 13
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 135
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1FECDES", 14)
'                        oColumn.Name = "BPE1FECDES"
'                        oColumn.Ordinal = 14
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 135
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1ESTPLA", 15)
'                        oColumn.Name = "BPE1ESTPLA"
'                        oColumn.Ordinal = 15
'                        oColumn.Flags = 120
'                        oColumn.Size = 9
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1ESTGEN", 16)
'                        oColumn.Name = "BPE1ESTGEN"
'                        oColumn.Ordinal = 16
'                        oColumn.Flags = 120
'                        oColumn.Size = 9
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1TIPAPI", 17)
'                        oColumn.Name = "BPE1TIPAPI"
'                        oColumn.Ordinal = 17
'                        oColumn.Flags = 120
'                        oColumn.Size = 1
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1TIPTEL", 18)
'                        oColumn.Name = "BPE1TIPTEL"
'                        oColumn.Ordinal = 18
'                        oColumn.Flags = 120
'                        oColumn.Size = 1
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1TIECOR", 19)
'                        oColumn.Name = "BPE1TIECOR"
'                        oColumn.Ordinal = 19
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 3
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1FECNEG", 20)
'                        oColumn.Name = "BPE1FECNEG"
'                        oColumn.Ordinal = 20
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 135
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1TIEBOR", 21)
'                        oColumn.Name = "BPE1TIEBOR"
'                        oColumn.Ordinal = 21
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 3
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1ESTCLI", 22)
'                        oColumn.Name = "BPE1ESTCLI"
'                        oColumn.Ordinal = 22
'                        oColumn.Flags = 120
'                        oColumn.Size = 10
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1PEDRET", 23)
'                        oColumn.Name = "BPE1PEDRET"
'                        oColumn.Ordinal = 23
'                        oColumn.Flags = 120
'                        oColumn.Size = 1
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1PRIORI", 24)
'                        oColumn.Name = "BPE1PRIORI"
'                        oColumn.Ordinal = 24
'                        oColumn.Flags = 120
'                        oColumn.Size = 1
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1FECOBJ", 25)
'                        oColumn.Name = "BPE1FECOBJ"
'                        oColumn.Ordinal = 25
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 135
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1DIVCLI", 26)
'                        oColumn.Name = "BPE1DIVCLI"
'                        oColumn.Ordinal = 26
'                        oColumn.Flags = 120
'                        oColumn.Size = 3
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1FECMOD", 27)
'                        oColumn.Name = "BPE1FECMOD"
'                        oColumn.Ordinal = 27
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 135
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1TIPERE", 28)
'                        oColumn.Name = "BPE1TIPERE"
'                        oColumn.Ordinal = 28
'                        oColumn.Flags = 120
'                        oColumn.Size = 1
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1CODERE", 29)
'                        oColumn.Name = "BPE1CODERE"
'                        oColumn.Ordinal = 29
'                        oColumn.Flags = 120
'                        oColumn.Size = 6
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1TIPIFI", 30)
'                        oColumn.Name = "BPE1TIPIFI"
'                        oColumn.Ordinal = 30
'                        oColumn.Flags = 120
'                        oColumn.Size = 1
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1CODACT", 31)
'                        oColumn.Name = "BPE1CODACT"
'                        oColumn.Ordinal = 31
'                        oColumn.Flags = 120
'                        oColumn.Size = 2
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1FECRET", 32)
'                        oColumn.Name = "BPE1FECRET"
'                        oColumn.Ordinal = 32
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 135
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1PURORD", 33)
'                        oColumn.Name = "BPE1PURORD"
'                        oColumn.Ordinal = 33
'                        oColumn.Flags = 120
'                        oColumn.Size = 15
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1CODDES", 34)
'                        oColumn.Name = "BPE1CODDES"
'                        oColumn.Ordinal = 34
'                        oColumn.Flags = 120
'                        oColumn.Size = 3
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1FECCAN", 35)
'                        oColumn.Name = "BPE1FECCAN"
'                        oColumn.Ordinal = 35
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 135
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1FEULDE", 36)
'                        oColumn.Name = "BPE1FEULDE"
'                        oColumn.Ordinal = 36
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 135
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1FECLIQ", 37)
'                        oColumn.Name = "BPE1FECLIQ"
'                        oColumn.Ordinal = 37
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 135
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1FECEMS", 38)
'                        oColumn.Name = "BPE1FECEMS"
'                        oColumn.Ordinal = 38
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 135
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1MARCAS", 39)
'                        oColumn.Name = "BPE1MARCAS"
'                        oColumn.Ordinal = 39
'                        oColumn.Flags = 120
'                        oColumn.Size = 1
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.SourceColumns.New("BPE1CODVAR", 40)
'                        oColumn.Name = "BPE1CODVAR"
'                        oColumn.Ordinal = 40
'                        oColumn.Flags = 120
'                        oColumn.Size = 1
'                        oColumn.DataType = 129
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.SourceColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1TIPPED", 1)
'                        oColumn.Name = "BPE1TIPPED"
'                        oColumn.Ordinal = 1
'                        oColumn.Flags = 120
'                        oColumn.Size = 1
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1CODPED", 2)
'                        oColumn.Name = "BPE1CODPED"
'                        oColumn.Ordinal = 2
'                        oColumn.Flags = 120
'                        oColumn.Size = 5
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1CODCLI", 3)
'                        oColumn.Name = "BPE1CODCLI"
'                        oColumn.Ordinal = 3
'                        oColumn.Flags = 120
'                        oColumn.Size = 5
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1CODEST", 4)
'                        oColumn.Name = "BPE1CODEST"
'                        oColumn.Ordinal = 4
'                        oColumn.Flags = 120
'                        oColumn.Size = 9
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1TIPTAR", 5)
'                        oColumn.Name = "BPE1TIPTAR"
'                        oColumn.Ordinal = 5
'                        oColumn.Flags = 120
'                        oColumn.Size = 1
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1CODTAR", 6)
'                        oColumn.Name = "BPE1CODTAR"
'                        oColumn.Ordinal = 6
'                        oColumn.Flags = 120
'                        oColumn.Size = 5
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1CLATAR", 7)
'                        oColumn.Name = "BPE1CLATAR"
'                        oColumn.Ordinal = 7
'                        oColumn.Flags = 120
'                        oColumn.Size = 1
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1CODPTA", 8)
'                        oColumn.Name = "BPE1CODPTA"
'                        oColumn.Ordinal = 8
'                        oColumn.Flags = 120
'                        oColumn.Size = 2
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1ULTOPR", 9)
'                        oColumn.Name = "BPE1ULTOPR"
'                        oColumn.Ordinal = 9
'                        oColumn.Flags = 120
'                        oColumn.Size = 3
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1UNIPED", 10)
'                        oColumn.Name = "BPE1UNIPED"
'                        oColumn.Ordinal = 10
'                        oColumn.Flags = 120
'                        oColumn.Size = 3
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1DESPED", 11)
'                        oColumn.Name = "BPE1DESPED"
'                        oColumn.Ordinal = 11
'                        oColumn.Flags = 120
'                        oColumn.Size = 35
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1INISAL", 12)
'                        oColumn.Name = "BPE1INISAL"
'                        oColumn.Ordinal = 12
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 7
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1FINSAL", 13)
'                        oColumn.Name = "BPE1FINSAL"
'                        oColumn.Ordinal = 13
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 7
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1FECDES", 14)
'                        oColumn.Name = "BPE1FECDES"
'                        oColumn.Ordinal = 14
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 7
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1ESTPLA", 15)
'                        oColumn.Name = "BPE1ESTPLA"
'                        oColumn.Ordinal = 15
'                        oColumn.Flags = 120
'                        oColumn.Size = 9
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1ESTGEN", 16)
'                        oColumn.Name = "BPE1ESTGEN"
'                        oColumn.Ordinal = 16
'                        oColumn.Flags = 120
'                        oColumn.Size = 9
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1TIPAPI", 17)
'                        oColumn.Name = "BPE1TIPAPI"
'                        oColumn.Ordinal = 17
'                        oColumn.Flags = 120
'                        oColumn.Size = 1
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1TIPTEL", 18)
'                        oColumn.Name = "BPE1TIPTEL"
'                        oColumn.Ordinal = 18
'                        oColumn.Flags = 120
'                        oColumn.Size = 1
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1TIECOR", 19)
'                        oColumn.Name = "BPE1TIECOR"
'                        oColumn.Ordinal = 19
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 3
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1FECNEG", 20)
'                        oColumn.Name = "BPE1FECNEG"
'                        oColumn.Ordinal = 20
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 7
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1TIEBOR", 21)
'                        oColumn.Name = "BPE1TIEBOR"
'                        oColumn.Ordinal = 21
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 3
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1ESTCLI", 22)
'                        oColumn.Name = "BPE1ESTCLI"
'                        oColumn.Ordinal = 22
'                        oColumn.Flags = 120
'                        oColumn.Size = 10
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1PEDRET", 23)
'                        oColumn.Name = "BPE1PEDRET"
'                        oColumn.Ordinal = 23
'                        oColumn.Flags = 120
'                        oColumn.Size = 1
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1PRIORI", 24)
'                        oColumn.Name = "BPE1PRIORI"
'                        oColumn.Ordinal = 24
'                        oColumn.Flags = 120
'                        oColumn.Size = 1
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1FECOBJ", 25)
'                        oColumn.Name = "BPE1FECOBJ"
'                        oColumn.Ordinal = 25
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 7
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1DIVCLI", 26)
'                        oColumn.Name = "BPE1DIVCLI"
'                        oColumn.Ordinal = 26
'                        oColumn.Flags = 120
'                        oColumn.Size = 3
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1FECMOD", 27)
'                        oColumn.Name = "BPE1FECMOD"
'                        oColumn.Ordinal = 27
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 7
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1TIPERE", 28)
'                        oColumn.Name = "BPE1TIPERE"
'                        oColumn.Ordinal = 28
'                        oColumn.Flags = 120
'                        oColumn.Size = 1
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1CODERE", 29)
'                        oColumn.Name = "BPE1CODERE"
'                        oColumn.Ordinal = 29
'                        oColumn.Flags = 120
'                        oColumn.Size = 6
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1TIPIFI", 30)
'                        oColumn.Name = "BPE1TIPIFI"
'                        oColumn.Ordinal = 30
'                        oColumn.Flags = 120
'                        oColumn.Size = 1
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1CODACT", 31)
'                        oColumn.Name = "BPE1CODACT"
'                        oColumn.Ordinal = 31
'                        oColumn.Flags = 120
'                        oColumn.Size = 2
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1FECRET", 32)
'                        oColumn.Name = "BPE1FECRET"
'                        oColumn.Ordinal = 32
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 7
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1PURORD", 33)
'                        oColumn.Name = "BPE1PURORD"
'                        oColumn.Ordinal = 33
'                        oColumn.Flags = 120
'                        oColumn.Size = 15
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1CODDES", 34)
'                        oColumn.Name = "BPE1CODDES"
'                        oColumn.Ordinal = 34
'                        oColumn.Flags = 120
'                        oColumn.Size = 3
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1FECCAN", 35)
'                        oColumn.Name = "BPE1FECCAN"
'                        oColumn.Ordinal = 35
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 7
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1FEULDE", 36)
'                        oColumn.Name = "BPE1FEULDE"
'                        oColumn.Ordinal = 36
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 7
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1FECLIQ", 37)
'                        oColumn.Name = "BPE1FECLIQ"
'                        oColumn.Ordinal = 37
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 7
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1FECEMS", 38)
'                        oColumn.Name = "BPE1FECEMS"
'                        oColumn.Ordinal = 38
'                        oColumn.Flags = 120
'                        oColumn.Size = 0
'                        oColumn.DataType = 7
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1MARCAS", 39)
'                        oColumn.Name = "BPE1MARCAS"
'                        oColumn.Ordinal = 39
'                        oColumn.Flags = 120
'                        oColumn.Size = 1
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'                Set oColumn = oTransformation.DestinationColumns.New("BPE1CODVAR", 40)
'                        oColumn.Name = "BPE1CODVAR"
'                        oColumn.Ordinal = 40
'                        oColumn.Flags = 120
'                        oColumn.Size = 1
'                        oColumn.DataType = 130
'                        oColumn.Precision = 0
'                        oColumn.NumericScale = 0
'                        oColumn.Nullable = True
'
'                oTransformation.DestinationColumns.Add oColumn
'                Set oColumn = Nothing
'
'        Set oTransProps = oTransformation.TransformServerProperties
'
'
'        Set oTransProps = Nothing
'
'        oCustomTask8.Transformations.Add oTransformation
'        Set oTransformation = Nothing
'
'End Sub
'
'
'
'
'
