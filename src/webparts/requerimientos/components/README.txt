# SharePoint Admin v1

Copiar en:

src/webparts/requerimientos/components/AdminSharePoint.tsx
src/webparts/requerimientos/components/admin/SharePointAdminModels.ts
src/webparts/requerimientos/components/admin/SharePointAdminService.ts

Tu RequerimientosWebPart.ts actual puede quedarse como está:
- URL normal => Requerimientos.tsx
- &admin=1 => AdminSharePoint.tsx

Incluye:
- Descubrimiento automático de listas
- Descubrimiento automático de campos
- Lectura de items
- Crear / editar / eliminar
- User y Lookup simples
- Choice, Boolean, Text, Note, Number, DateTime
- Selección y edición masiva
- REST Console con confirmación para escritura
- Vista de InternalName y TypeAsString

Notas v1:
- Se muestran hasta 200 elementos por lista en el explorador.
- MultiChoice, UserMulti y LookupMulti todavía requieren soporte específico.
- Adjuntos y bibliotecas de documentos no forman parte del CRUD de esta v1.
- La seguridad real sigue dependiendo de permisos SharePoint. ?admin=1 solo controla la UI.
