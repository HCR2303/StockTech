Attribute VB_Name = "TablasDB"
Option Private Module
Const MODULE_NAME As String = "TablasDB"
Public Enum CNameSolicitudes
    sol_id_solicitud = 0
    sol_fecha = 1
    sol_usuario = 2
    sol_codigo = 3
    sol_tipo = 4
    sol_unidades = 5
    sol_realizada = 6
    sol_comentarios = 7
End Enum

Public Enum CNameUsuarios
    us_id = 0
    us_nombre = 1
    us_usuario = 2
    us_correo = 3
    us_nivel = 4
    us_contrasenia = 5
End Enum

Public Enum CNamePrivilegios
    priv_Id = 0
    priv_nivel = 1
    priv_NewCodsBttn = 2
    priv_NewSolicitudBttn = 3
    priv_ConsultDBBttn = 4
    priv_PresupuestosBttn = 5
    priv_SolicitudBttn = 6
    priv_SeguimientoBttn = 7
    priv_AnalisisUserBttn = 8
    priv_VerTablaBttn = 9
    priv_UpdateBttn = 10
    priv_EliminarBttn = 11
    priv_EditarBttn = 12
    priv_usersBttn = 13
    priv_AuditBttn = 14
    priv_salirBttn = 15
End Enum

Public Enum CNamePresupuestos
    pre_id = 0
    pre_centro_costo = 1
    pre_id_cuenta = 2
    pre_cuenta_anterior = 3
    pre_presupuesto = 4
End Enum

Public Enum CNameRefacciones
    ref_id = 0
    ref_material = 1
    ref_descripcion = 2
    ref_presentacion = 3
    ref_stock = 4
    ref_ubicacion_almacen = 5
    ref_criticidad = 6
End Enum

Public Enum CNameProveedores
    prv_id = 0
    prv_codigo = 1
    prv_nombre = 2
End Enum

' TCuentas -> Prefijo: cta_
Public Enum CNameCuentas
    cta_id = 0
    cta_id_cuenta = 1
    cta_nombre = 2
End Enum

' TEquipos -> Prefijo: equ_
Public Enum CNameEquipos
    equ_id = 0
    equ_codigo = 1
    equ_nombre = 2
    equ_marca = 3
    equ_modelo = 4
    equ_anio_construccion = 5
    equ_no_serie = 6
    equ_cto_actualizacion = 7
    equ_criticidad = 8
    equ_id_ubicacion = 9
    equ_gpo_planificacion = 10
    equ_cto_planificacion = 11
    equ_cto_trabajo_mantenimiento = 12
    equ_cto_puesto_trabajo = 13
    equ_sociedad = 14
End Enum

Public Enum CNameUbicaciones
    ubi_id = 0
    ubi_id_ubicacion = 1
    ubi_indicador = 2
    ubi_categoria = 3
    ubi_ubicacion = 4
    ubi_gpo_autorizacion = 5
    ubi_cto_actualizacion = 6
    ubi_gpo_planificacion = 7
    ubi_centro_planificacion = 8
    ubi_cto_trabajo_mantenimiento = 9
    ubi_cto_puesto_trabajo = 10
    ubi_sociedad = 11
    ubi_centro_costo = 12
End Enum


Public Enum CNameSOLPEDs
    spd_id = 0
    SPD_SOLPED = 1
    spd_cuenta = 2
    spd_codigo_equipo = 3
    SPD_REFACCION = 4
    spd_proveedor = 5
    spd_proveedor_unico = 6
    spd_orden_compra = 7
    spd_no_factura = 8
    spd_unidades = 9
    spd_costo = 10
    spd_capturo = 11
    spd_id_solicitud = 12
    spd_uso = 13
End Enum


Public Enum CNameSOLPEDsTrack
    spt_id = 0
    spt_usuario = 1
    spt_fecha = 2
    spt_estado = 3
    spt_solped = 4
    spt_comentario = 5
End Enum


Public Enum CNameAuditTrail
    aud_id = 0
    aud_usuario = 1
    aud_fecha = 2
    aud_accion = 3
    aud_objetivo = 4
    aud_detalles = 5
    aud_comentario = 6
End Enum


Public Enum CNameSystemTrack
    sys_id = 0
    sys_fecha = 1
    sys_usuario = 2
    sys_lugar = 3
    sys_error = 4
End Enum

Public Enum CNameAccesos
    acc_id = 0
    acc_tipo = 1
    acc_contrasenia = 2
End Enum

'=========================================================================
'   En este módulo se declaran los nombres generalizados para la base
'   datos con a que el Programador relacionará Tablas de Access con
'   Tablas de excel
'=========================================================================

' TSolicitudes: tabla de registro de solicitud de SOLPED
Public Const TSolicitudes As String = "Solicitudes"
Public Const FieldsSolicitudes As String = "Id_Solicitud|Fecha|Usuario|Código|Tipo|Unidades|Realizada|Comentarios"
' TPresupuestos: Tabla de presupuestos anuales cargada por el usuario
Public Const TPresupuestos As String = "Presupuestos"
Public Const FieldsPresupuestos As String = "Id|Centro de Costo|Id_Cuenta|Cuenta Anteriror|Presupuesto"

'    ---> Tablas SAP <---
Public Const TRefacciones As String = "Refacciones" ' Descripción de refacciones por código
Public Const FieldsRefacciones As String = "Id|Material|Descripción|Presentación|Stock|UbicaciónAlmacén|Criticidad"

Public Const TProveedores As String = "Proveedores" ' Descripción de proveedores por código
Public Const FieldsProveedores As String = "Id|Código|Nombre"

Public Const TCuentas As String = "Cuentas HANNA" ' Decripción de cuentas por número
Public Const FieldsCuentas As String = "Id|Id_Cuenta|Nombre"

Public Const TEquipos As String = "Equipos" ' Descripción de equipos por código
Public Const FieldsEquipos As String = "Id|Código|Nombre|Marca|Modelo|Año Construcción|No de Serie|Cto Actualización|Criticidad|Id_Ubicación|" & _
                                        "Gpo Planificación|Cto de Planificación|Cto de Trabajo Mantenimiento|Cto de Puesto Trabajo|Sociedad"

Public Const TUbicaciones As String = "Ubicaciones" ' Descripción de ubicaciones por código
Public Const FieldsUbicaciones As String = "Id|Id_Ubicación|Indicador|Categoría|Ubicación|Gpo Autorización|Cto de Actualización|Gpo Planificación|" & _
                                            "Centro de Planificación|Cto Trabajo Mantenimiento|Cto Puesto de Trabajo|Sociedad|Centro de Costo"

' TSOLPEDs: Tabla de registro de SOLPEDs creadas con descripción por códigos
Public Const TSOLPEDs As String = "SOLPEDs"
Public Const FieldsSOLPEDs As String = "Id|SOLPED|Cuenta|Codigo de Equipo|Refacción|Proveedor|Proveedor Único|" & _
                                        "Orden de Compra|No_Factura|Unidades|Costo|Capturo|Id_Solicitud|Uso"

Public Const TUsuarios As String = "Usuarios"
Public Const FieldsUsuarios As String = "Id|Nombre|Usuario|Correo|Nivel|Contraseña"

Public Const TPrivilegios As String = "Privilegios"
Public Const FieldsPrivilegios As String = "Id|Nivel|NewCodsBttn|NewSolicitudBttn|ConsultDBBttn|PresupuestosBttn|SolicitudBttn|" & _
                                           "SeguimientoBttn|AnalisisUserBttn|VerTablaBttn|UpdateBttn|EliminarBttn|EditarBttn|usersBttn|" & _
                                           "AuditBttn|salirBttn"


' TSOLPEDsTrack: Tabla de cambios de estado de SOLPEDs
Public Const TSOLPEDsTrack As String = "Seguimiento SOLPEDs"
Public Const FieldsSOLPEDsTrack As String = "Id|Usuario|Fecha|Estado|SOLPED|Comentario"

Public Const TAuditTrail As String = "AuditTrail"
Public Const FieldsAuditTrail As String = "Id|Usuario|Fecha|Acción|Objetivo|Detalles|Comentario"

Public Const TSystemTrack As String = "ErrorsTrack"
Public Const FieldsSystemTrack As String = "Id|Fecha|Usuario|Lugar|Error"

Public Const TAccesos As String = "Accesos"
Public Const FieldsAccesos As String = "Id|Tipo|Contraseña"

Public Function CamposTablaDB(ByVal tabla As String) As Variant
    Select Case tabla
        Case TSolicitudes
            CamposTablaDB = Split(FieldsSolicitudes, "|")
        Case TPresupuestos
            CamposTablaDB = Split(FieldsPresupuestos, "|")
        Case TRefacciones
            CamposTablaDB = Split(FieldsRefacciones, "|")
        Case TProveedores
            CamposTablaDB = Split(FieldsProveedores, "|")
        Case TCuentas
            CamposTablaDB = Split(FieldsCuentas, "|")
        Case TEquipos
            CamposTablaDB = Split(FieldsEquipos, "|")
        Case TUbicaciones
            CamposTablaDB = Split(FieldsUbicaciones, "|")
        Case TSOLPEDs
            CamposTablaDB = Split(FieldsSOLPEDs, "|")
        Case TSOLPEDsTrack
            CamposTablaDB = Split(FieldsSOLPEDsTrack, "|")
        Case TAuditTrail
            CamposTablaDB = Split(FieldsAuditTrail, "|")
        Case TSystemTrack
            CamposTablaDB = Split(FieldsSystemTrack, "|")
        Case TUsuarios
            CamposTablaDB = Split(FieldsUsuarios, "|")
        Case TPrivilegios
            CamposTablaDB = Split(FieldsPrivilegios, "|")
        Case TAccesos
            CamposTablaDB = Split(FieldsAccesos, "|")
        Case Else
            CamposTablaDB = Empty
    End Select
End Function
