<div align="center">
  <h1>📦 StockTech</h1>
  <p><i>Sistema multiusuario corporativo para la gestión de compras y administración de almacén.</i></p>

  <img src="https://img.shields.io/badge/VBA-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white" alt="VBA" />
  <img src="https://img.shields.io/badge/Microsoft_Access-A4373A?style=for-the-badge&logo=microsoftaccess&logoColor=white" alt="Access" />
  <img src="https://img.shields.io/badge/ADODB-4479A1?style=for-the-badge" alt="ADODB" />
</div>

---

## 📝 Descripción del Proyecto

**StockTech** es una solución de software a medida diseñada para centralizar y optimizar la gestión de adquisiciones y el control de inventario en las instalaciones de Toluca. Implementando una arquitectura cliente-servidor de forma local, el aplicativo utiliza interfaces desarrolladas en VBA que se conectan de manera transaccional a una base de datos relacional centralizada en Microsoft Access, permitiendo la operación simultánea de múltiples usuarios.

## ✨ Características Principales

* **Concurrencia Multiusuario:** Gestión de conexiones simultáneas mediante tecnología ADODB, previniendo bloqueos de registros durante la edición de inventarios.
* **Control de Jerarquías:** Sistema de privilegios basado en roles para delimitar las operaciones permitidas según el perfil del usuario (Operadores de almacén, Químicos de laboratorio, Supervisores y Directores).
* **Trazabilidad de Transacciones:** Registro histórico de entradas, salidas y órdenes de compra para garantizar la auditabilidad de los flujos de trabajo.
* **Interfaces Optimizadas:** Formularios interactivos que reemplazan las bitácoras físicas, reduciendo la carga administrativa y minimizando errores de captura.

## 🏗️ Arquitectura del Sistema

* **Frontend:** Formularios y módulos lógicos programados en VBA. Contiene la lógica de validación de datos y la presentación al usuario final.
* **Backend:** Archivo centralizado `.accdb` (Microsoft Access) alojado en un servidor o unidad de red compartida. Contiene exclusivamente el esquema de tablas relacionales y las consultas estructuradas (SQL).
* **Capa de Conexión:** Cadenas de conexión OLEDB/ADODB para enviar y recuperar instrucciones SQL de manera segura y eficiente.

## ⚙️ Prerrequisitos y Despliegue

Para la correcta ejecución de este aplicativo en estaciones de trabajo corporativas:
1. **Entorno:** Microsoft Office instalado (versión 2016 o superior).
2. **Conectividad:** Acceso de lectura y escritura a la unidad de red compartida donde reside el archivo backend de Access.
3. **Seguridad:** Habilitar la ejecución de macros en el Centro de Confianza de la aplicación cliente para permitir la inicialización del código VBA.

## 🔒 Auditoría y Calidad

El código está estructurado bajo principios de formalidad profesional, asegurando el cierre explícito de conexiones a la base de datos (`Recordset.Close` y `Connection.Close`) y la liberación de variables de objeto en memoria al finalizar cada transacción. Esto previene la corrupción de la base de datos y mantiene la estabilidad del aplicativo en entornos de producción intensiva.
