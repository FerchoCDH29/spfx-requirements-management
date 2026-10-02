## Versionado

El proyecto utiliza versionado semántico `MAJOR.MINOR.PATCH`.

- `npm run version:patch` → correcciones.
- `npm run version:minor` → nuevas funcionalidades.
- `npm run version:major` → cambios importantes.
- `npm run build` → genera el paquete `.sppkg`.

La versión de SPFx se sincroniza automáticamente con `package-solution.json`.

# SPFx Requirements Management

Aplicación desarrollada con **SharePoint Framework (SPFx)** y **React** para la gestión de requerimientos internos dentro de SharePoint.

El proyecto permite registrar solicitudes, seleccionar categorías, adjuntar documentos y preparar el flujo de información para procesos de revisión y aprobación.

## Tecnologías utilizadas

* SharePoint Framework (SPFx)
* React
* TypeScript
* SCSS
* Microsoft 365
* SharePoint Online
* Power Automate

## Funcionalidades principales

Actualmente el Web Part permite:

* Registrar nuevos requerimientos.
* Obtener información del usuario autenticado.
* Mostrar datos del solicitante.
* Seleccionar una categoría de requerimiento.
* Indicar si una solicitud es recurrente.
* Registrar el valor total del requerimiento.
* Adjuntar documentos PDF.
* Crear el requerimiento dentro de SharePoint.
* Preparar la información para procesos de aprobación.

## Estructura principal

```text
spfx-requirements-management/
│
├── config/
├── src/
│   └── webparts/
│       └── requerimientos/
│           ├── assets/
│           ├── components/
│           │   ├── IRequerimientosProps.ts
│           │   ├── Requerimientos.module.scss
│           │   └── Requerimientos.tsx
│           ├── loc/
│           ├── RequerimientosWebPart.manifest.json
│           └── RequerimientosWebPart.ts
│
├── teams/
├── package.json
├── package-lock.json
├── tsconfig.json
└── README.md
```

## Requisitos

Para ejecutar el proyecto localmente se necesita:

* Node.js compatible con la versión de SPFx utilizada.
* npm
* SharePoint Online
* Acceso a un tenant de Microsoft 365.
* Certificado de desarrollo de SPFx configurado.

## Instalación

Clonar el repositorio:

```bash
git clone https://github.com/FerchoCDH29/spfx-requirements-management.git
```

Ingresar al proyecto:

```bash
cd spfx-requirements-management
```

Instalar dependencias:

```bash
npm install
```

## Ejecutar en desarrollo

Confiar en el certificado de desarrollo:

```bash
heft trust-dev-cert
```

Ejecutar el proyecto:

```bash
heft start
```

Después se puede utilizar el Workbench de SharePoint para probar el Web Part.

## Compilar el proyecto

Para generar una compilación de producción:

```bash
heft build --production
```

## Generar paquete SPFx

El proyecto utiliza la configuración definida dentro de:

```text
config/package-solution.json
```

El paquete generado puede ser utilizado posteriormente para desplegar la solución en el App Catalog de SharePoint.

## Flujo general

```text
Usuario
   ↓
SPFx Web Part
   ↓
Formulario de requerimiento
   ↓
SharePoint
   ↓
Proceso de aprobación
   ↓
Power Automate
```

## Estado del proyecto

Proyecto actualmente en desarrollo.

Se están implementando y mejorando funcionalidades relacionadas con:

* Gestión de requerimientos.
* Experiencia de usuario.
* Diseño del formulario.
* Manejo de documentos adjuntos.
* Integración con SharePoint.
* Flujos de aprobación.
* Gestión de usuarios, roles y aprobadores.

## Repositorio

GitHub:

```text
https://github.com/FerchoCDH29/spfx-requirements-management
```

## Autor

**fcardenas**

Proyecto desarrollado como parte de una solución de gestión de requerimientos utilizando tecnologías del ecosistema Microsoft 365.
