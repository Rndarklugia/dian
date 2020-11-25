Imports System.Data
Imports System.Data.OleDb
Module OperacionesBD
    Dim MiConexion As New OleDbConnection
    Dim MiComando As New OleDbCommand
    Dim Estado As String
    Public Sub Conectar()
        Try
            MiConexion.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\wjeff\source\repos\WindowsApp1\Personas.mdb"
            MiConexion.Open()
            Estado = "Conectado"
            'MsgBox(Estado, MsgBoxStyle.OkOnly, "Estado de la Conexión")
            Form1.lblEstadoConexion.Text = Estado
            MostrarDatos()
        Catch ex As Exception
            Estado = "Sin Conectar"
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Estado de la Conexión")
            Form1.lblEstadoConexion.Text = Estado
            MiConexion.Close()
        End Try
    End Sub

    Public Sub MostrarDatos()
        Dim oda As New OleDbDataAdapter
        Dim ods As New DataSet
        Dim MyQuery As String

        MyQuery = "Select * From Jugadores ORDER By Id"
        oda = New OleDbDataAdapter(MyQuery, MiConexion)
        ods.Tables.Add("TablaJugadores")
        oda.Fill(ods.Tables("TablaJugadores"))
        Form1.DataGridView1.DataSource = ods.Tables("TablaJugadores")

    End Sub
    Public Function Insertar(_Nombres As String, _Apellidos As String, _Edad As String)

        Dim MiInstruccionSQL As String
        Try
            'MiConexion.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Roy\Programación II\Proyectos de Clase\Acceso a Datos\Personas.mdb"
            'MiConexion.Open()

            'MsgBox(Estado, MsgBoxStyle.OkOnly, "Estado de la Conexión")

            MiInstruccionSQL = "INSERT INTO Jugadores (Nombre, Apellidos, Edad, idEquipo) VALUES ( '" &
                                     _Nombres & "', '" & _Apellidos & "', " & _Edad & ",1)"
            MiComando = New OleDbCommand(MiInstruccionSQL, MiConexion)
            MiComando.ExecuteNonQuery()
            'MiConexion.Close()
            Estado = "Datos Guardados Correctamente"
            MostrarDatos()
            Return Estado
        Catch ex As Exception
            Estado = "Error al tratar de Guardar registro"
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Estado de la Conexión")
            MiConexion.Close()
            Return Estado
        End Try



    End Function

    Public Function Desconectar()
        Try
            MiConexion.Close()
            Estado = "Desconectado"
            MsgBox(Estado, MsgBoxStyle.OkOnly, "Estado de la Conexión")

        Catch ex As Exception
            Estado = "Sin Desonectar"

            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Estado de la Conexión")
            Return Estado
        End Try
    End Function

End Module
