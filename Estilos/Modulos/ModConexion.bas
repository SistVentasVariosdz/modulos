Attribute VB_Name = "ModConexion"
Public Sub Realiza_Conexion()
    Set B_db = Nothing
    B_db.ConnectionString = cCONNECT
    B_db.Open
End Sub


