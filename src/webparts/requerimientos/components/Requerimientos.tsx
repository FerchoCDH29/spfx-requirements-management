import * as React from 'react';
import './global/global.module.scss';
import { useEffect, useState } from 'react';
import { SPHttpClient } from '@microsoft/sp-http';

import styles from './Requerimientos.module.scss';
import inovaLogo from '../assets/inova-logo.png';
import { IRequerimientosProps } from './IRequerimientosProps';

const MODO_PRUEBAS = true;

type UserRole =
  | 'Admin'
  | 'Solicitador'
  | 'Aprobador'
  | 'ResponsableOC'
  | 'SinRol';

type AppView =
  | 'nueva'
  | 'mis'
  | 'pendientes'
  | 'aprobadas'
  | 'rechazadas'
  | 'todas'
  | 'ordenes';

interface IUser {
  id: number;
  displayName: string;
  email: string;
  role: UserRole;
  areaId?: number;
  area?: string;
  roleId?: number;
}

interface ISharePointUser {
  Id: number;
  Title: string;
  Email: string;
}

interface ICategory {
  Id: number;
  NombreCategoria: string;
  AprobadorCategoriaId?: number;
  AprobadorPruebaId?: number;
}

interface IAdjuntoCotizacion {
  FileName: string;
  ServerRelativeUrl: string;
}

interface ICotizacionForm {
  key: string;
  id?: number;
  valorTotal: string;
  archivo: File | null;
  adjuntosExistentes: IAdjuntoCotizacion[];
}

interface IRequerimientoItem {
  Id: number;
  Title: string;
  Descripcion: string;
  Area?: string;
  CategoriaId?: number;
  AprobadorId?: number;
  ValorPromedio?: number | null;
  ValorTotal?: number | null;
  Estado: string;
  Recurrente: boolean;
  Created: string;
  EtapaActual?: string;
  FechaAprobacion?: string;
  FechaEnvioAprobacion?: string;
  FechaDecision?: string;
  ComentarioAprobador?: string;

  Aprobador?: {
    Id: number;
    Title: string;
    EMail?: string;
  };

  DecisionPor?: {
    Id: number;
    Title: string;
    EMail?: string;
  };

  Solicitante?: {
    Id: number;
    Title: string;
    EMail?: string;
  };
}

interface ICotizacionDetalle {
  Id: number;
  Title: string;
  NumeroCotizacionProveedor?: string;
  CodigoCotizacionInterno?: string;
  ValorTotal: number;
  Seleccionada: boolean;
  Attachments: IAdjuntoCotizacion[];
}

interface IResponsableOC {
  Id: number;
  Title: string;
  EMail?: string;
}

interface IOrdenCompra {
  Id: number;
  Title: string;
  RequerimientoId?: number;
  CotizacionId?: number;
  TipoOrden: 'GenerarOC' | 'SolicitarOC' | string;
  TipoDocumento: 'Contrato' | 'OCV' | string;
  EstadoOC: 'Pendiente' | 'En proceso' | 'Generada' | 'Rechazada' | string;
  NumeroOC?: string | null;
  FechaSolicitud?: string;
  FechaGeneracion?: string | null;
  Modified?: string;
  Observaciones?: string | null;
  AttachmentFiles: IAdjuntoCotizacion[];
  ResponsableOCId?: number;
  SolicitadoPor?: {
    Id: number;
    Title: string;
    EMail?: string;
  };
  ResponsableOC?: {
    Id: number;
    Title: string;
    EMail?: string;
  };
}

interface IHistoryEvent {
  key: string;
  title: string;
  date?: string;
  actor?: string;
  detail?: string;
}

interface IStageInspection {
  title: string;
  status: string;
  date?: string;
  actor?: string;
  details: string[];
}

const Requerimientos: React.FC<IRequerimientosProps> = ({
  context
}) => {

  /*
   * ===========================
   * ESTADO GENERAL
   * ===========================
   */

  const [user, setUser] =
    useState<IUser | null>(null);

  const [categories, setCategories] =
    useState<ICategory[]>([]);

  const [activeView, setActiveView] =
    useState<AppView>('nueva');

  const [requerimientos, setRequerimientos] =
    useState<IRequerimientoItem[]>([]);

  const [ordenesCompra, setOrdenesCompra] =
    useState<IOrdenCompra[]>([]);

  const [loadingOrdenesCompra, setLoadingOrdenesCompra] =
    useState<boolean>(false);

  const [selectedOrdenCompra, setSelectedOrdenCompra] =
    useState<IOrdenCompra | null>(null);

  const [ordenCompraDetalle, setOrdenCompraDetalle] =
    useState<IOrdenCompra | null>(null);

  const [numeroOCInput, setNumeroOCInput] =
    useState<string>('');

  const [observacionesOCInput, setObservacionesOCInput] =
    useState<string>('');

  const [procesandoOrdenCompra, setProcesandoOrdenCompra] =
    useState<boolean>(false);

  const [facturaPdf, setFacturaPdf] =
    useState<File | null>(null);

  const [procesandoFactura, setProcesandoFactura] =
    useState<boolean>(false);

  const [loading, setLoading] =
    useState<boolean>(true);

  const [
    loadingRequerimientos,
    setLoadingRequerimientos
  ] = useState<boolean>(false);

  const [saving, setSaving] =
    useState<boolean>(false);

  const [message, setMessage] =
    useState<string>('');

  const [selectedRequerimiento, setSelectedRequerimiento] =
    useState<IRequerimientoItem | null>(null);

  const [cotizacionesDetalle, setCotizacionesDetalle] =
    useState<ICotizacionDetalle[]>([]);

  const [loadingDetalle, setLoadingDetalle] =
    useState<boolean>(false);

  const [comentarioAprobador, setComentarioAprobador] = useState<string>('');
  const [procesandoDecision, setProcesandoDecision] = useState<boolean>(false);
  const [cotizacionSeleccionadaId, setCotizacionSeleccionadaId] =
    useState<number | null>(null);

  const [archivoPreview, setArchivoPreview] =
    useState<IAdjuntoCotizacion | null>(null);

  const [searchTerm, setSearchTerm] =
    useState<string>('');

  const [etapaConsulta, setEtapaConsulta] =
    useState<string | null>(null);

  /*
   * ===========================
   * ESTADO DEL FORMULARIO
   * ===========================
   */

  const [descripcion, setDescripcion] =
    useState<string>('');

  const [categoriaId, setCategoriaId] =
    useState<string>('');

  const [recurrente, setRecurrente] =
    useState<boolean>(false);

  const createEmptyCotizacion = (): ICotizacionForm => ({
    key: `${Date.now()}-${Math.random()}`,
    valorTotal: '',
    archivo: null,
    adjuntosExistentes: []
  });

  const [cotizaciones, setCotizaciones] =
    useState<ICotizacionForm[]>([createEmptyCotizacion()]);

  const [editingDraftId, setEditingDraftId] =
    useState<number | null>(null);

  const [originalDraftCotizacionIds, setOriginalDraftCotizacionIds] =
    useState<number[]>([]);

  /*
   * ===========================
   * UTILIDADES
   * ===========================
   */

  const getErrorMessage =
    (error: unknown): string => {

      if (error instanceof Error) {
        return error.message;
      }

      return 'Ocurrió un error inesperado.';
    };

  const getCategoriaNombre =
    (id?: number): string => {

      if (!id) {
        return 'Sin categoría';
      }

      const categoria =
        categories.find(
          (item: ICategory) =>
            item.Id === id
        );

      return (
        categoria?.NombreCategoria ||
        'Sin categoría'
      );
    };

  const getTabClass =
    (view: AppView): string => {

      return activeView === view
        ? `${styles.tabButton} ${styles.tabButtonActive}`
        : styles.tabButton;
    };

  /*
   * ===========================
   * CONFIGURACIÓN DEL USUARIO
   * FUENTE: LISTA UsuariosRoles
   * ===========================
   */

  const getCurrentUserConfiguration =
    async (currentUserId: number): Promise<{
      role: UserRole;
      areaId?: number;
      area?: string;
      roleId?: number;
    }> => {

      const usuariosRolesUrl =
        `${context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('UsuariosRoles')/items` +
        `?$select=Id,UsuarioId,Area0Id,RolId,Activo` +
        `&$filter=UsuarioId eq ${currentUserId} and Activo eq 1` +
        `&$top=1`;

      const response = await context.spHttpClient.get(
        usuariosRolesUrl,
        SPHttpClient.configurations.v1,
        { headers: { Accept: 'application/json;odata=nometadata' } }
      );

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(
          `No se pudo obtener la configuración del usuario. HTTP ${response.status}: ${errorText}`
        );
      }

      const data = await response.json();
      const usuarioRol = data.value?.[0] as {
        Area0Id?: number;
        RolId?: number;
        Activo?: boolean;
      } | undefined;

      if (!usuarioRol || !usuarioRol.Activo) {
        return { role: 'SinRol' };
      }

      let area: string | undefined;
      if (usuarioRol.Area0Id) {
        const areaUrl =
          `${context.pageContext.web.absoluteUrl}` +
          `/_api/web/lists/getbytitle('Areas')/items(${usuarioRol.Area0Id})` +
          `?$select=Id,NombreArea,Activo`;

        const areaResponse = await context.spHttpClient.get(
          areaUrl,
          SPHttpClient.configurations.v1,
          { headers: { Accept: 'application/json;odata=nometadata' } }
        );

        if (!areaResponse.ok) {
          const errorText = await areaResponse.text();
          throw new Error(
            `No se pudo obtener el área del usuario. HTTP ${areaResponse.status}: ${errorText}`
          );
        }

        const areaData = await areaResponse.json();
        if (areaData.Activo !== false) {
          area = areaData.NombreArea;
        }
      }

      let role: UserRole = 'SinRol';
      if (usuarioRol.RolId) {
        const rolUrl =
          `${context.pageContext.web.absoluteUrl}` +
          `/_api/web/lists/getbytitle('Rol')/items(${usuarioRol.RolId})` +
          `?$select=Id,nombreRol,Activo`;

        const rolResponse = await context.spHttpClient.get(
          rolUrl,
          SPHttpClient.configurations.v1,
          { headers: { Accept: 'application/json;odata=nometadata' } }
        );

        if (!rolResponse.ok) {
          const errorText = await rolResponse.text();
          throw new Error(
            `No se pudo obtener el rol del usuario. HTTP ${rolResponse.status}: ${errorText}`
          );
        }

        const rolData = await rolResponse.json();
        const nombreRol = String(rolData.nombreRol || '').trim().toLowerCase();

        if (rolData.Activo !== false) {
          if (nombreRol === 'admin' || nombreRol === 'administrador') {
            role = 'Admin';
          } else if (nombreRol === 'aprobador') {
            role = 'Aprobador';
          } else if (nombreRol === 'solicitador' || nombreRol === 'solicitante') {
            role = 'Solicitador';
          } else if (nombreRol === 'responsableoc') {
            role = 'ResponsableOC';
          }
        }
      }

      console.log('CONFIGURACIÓN USUARIO:', {
        role,
        areaId: usuarioRol.Area0Id,
        area,
        roleId: usuarioRol.RolId
      });

      return {
        role,
        areaId: usuarioRol.Area0Id,
        area,
        roleId: usuarioRol.RolId
      };
    };

  /*
   * ===========================
   * CATEGORÍAS
   * FUENTE: LISTA SHAREPOINT
   * ===========================
   */

  const loadCategories =
    async (areaId?: number, loadAll: boolean = false): Promise<void> => {

      if (!loadAll && !areaId) {
        setCategories([]);
        return;
      }

      let categoriesUrl =
        `${context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('Categorias')/items` +
        `?$select=Id,NombreCategoria,AreaId,Activo,AprobadorCategoriaId,AprobadorPruebaId` +
        `&$orderby=NombreCategoria asc` +
        `&$top=500`;

      const filter = loadAll
        ? `Activo eq 1`
        : `AreaId eq ${areaId} and Activo eq 1`;

      categoriesUrl += `&$filter=${encodeURIComponent(filter)}`;

      const response = await context.spHttpClient.get(
        categoriesUrl,
        SPHttpClient.configurations.v1,
        { headers: { Accept: 'application/json;odata=nometadata' } }
      );

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(
          `Error cargando categorías. HTTP ${response.status}: ${errorText}`
        );
      }

      const data = await response.json();

      const mappedCategories: ICategory[] = data.value.map(
            (item: {
              Id: number;
              NombreCategoria?: string;
              AprobadorCategoriaId?: number;
              AprobadorPruebaId?: number;
            }) => ({
              Id: item.Id,
              NombreCategoria: item.NombreCategoria || `Categoría ${item.Id}`,
              AprobadorCategoriaId: item.AprobadorCategoriaId,
              AprobadorPruebaId: item.AprobadorPruebaId
            })
          );
      console.log('CATEGORÍAS DEL ÁREA:', mappedCategories);
      setCategories(mappedCategories);
    };


  /*
   * ===========================
   * ÓRDENES DE COMPRA
   * FUENTE: LISTA SHAREPOINT
   * ===========================
   */

  const loadOrdenesCompra = async (
    responsableOCId: number
  ): Promise<void> => {
    try {
      setLoadingOrdenesCompra(true);
      setMessage('');

      const filter = encodeURIComponent(
        `ResponsableOCId eq ${responsableOCId} and Activo eq 1`
      );

      const ordenesUrl =
        `${context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('OrdenesCompra')/items` +
        `?$select=` +
        `Id,Title,RequerimientoId,CotizacionId,TipoOrden,TipoDocumento,EstadoOC,NumeroOC,` +
        `FechaSolicitud,FechaGeneracion,Modified,Observaciones,ResponsableOCId,` +
        `AttachmentFiles/FileName,AttachmentFiles/ServerRelativeUrl,` +
        `ResponsableOC/Id,ResponsableOC/Title,ResponsableOC/EMail,` +
        `SolicitadoPor/Id,SolicitadoPor/Title,SolicitadoPor/EMail` +
        `&$expand=ResponsableOC,SolicitadoPor,AttachmentFiles` +
        `&$filter=${filter}` +
        `&$orderby=FechaSolicitud desc` +
        `&$top=500`;

      const response = await context.spHttpClient.get(
        ordenesUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata'
          }
        }
      );

      if (!response.ok) {
        throw new Error(
          `Error cargando órdenes de compra. HTTP ${response.status}: ${await response.text()}`
        );
      }

      const data = await response.json();
      const ordenes = (data.value || []).map((orden: IOrdenCompra) => ({
        ...orden,
        AttachmentFiles: orden.AttachmentFiles || []
      })) as IOrdenCompra[];
      setOrdenesCompra(ordenes);
    } catch (error) {
      console.error('Error cargando órdenes de compra:', error);
      setMessage(getErrorMessage(error));
    } finally {
      setLoadingOrdenesCompra(false);
    }
  };



  const getOrdenCompraPorRequerimiento = async (
    requerimientoId: number
  ): Promise<IOrdenCompra | null> => {
    const filter = encodeURIComponent(
      `RequerimientoId eq ${requerimientoId} and Activo eq 1`
    );

    const ordenUrl =
      `${context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists/getbytitle('OrdenesCompra')/items` +
      `?$select=` +
      `Id,Title,RequerimientoId,CotizacionId,TipoOrden,TipoDocumento,EstadoOC,NumeroOC,` +
      `FechaSolicitud,FechaGeneracion,Modified,Observaciones,ResponsableOCId,` +
      `AttachmentFiles/FileName,AttachmentFiles/ServerRelativeUrl,` +
      `ResponsableOC/Id,ResponsableOC/Title,ResponsableOC/EMail,` +
      `SolicitadoPor/Id,SolicitadoPor/Title,SolicitadoPor/EMail` +
      `&$expand=ResponsableOC,SolicitadoPor,AttachmentFiles` +
      `&$filter=${filter}` +
      `&$orderby=Id desc&$top=1`;

    const response = await context.spHttpClient.get(
      ordenUrl,
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=nometadata' } }
    );

    if (!response.ok) {
      throw new Error(
        `No se pudo cargar la orden de compra del requerimiento #${requerimientoId}. ` +
        `HTTP ${response.status}: ${await response.text()}`
      );
    }

    const data = await response.json();

    if (!data.value?.length) {
      return null;
    }

    const orden = data.value[0] as IOrdenCompra;
    return {
      ...orden,
      AttachmentFiles: orden.AttachmentFiles || []
    };
  };


  /*
   * ===========================
   * REQUERIMIENTOS
   * FUENTE: LISTA SHAREPOINT
   * ===========================
   */

  const loadRequerimientos =
    async (
      view: AppView,
      userId: number
    ): Promise<void> => {

      if (view === 'nueva') {
        return;
      }

      try {

        setLoadingRequerimientos(true);
        setMessage('');

        let filter = '';

        switch (view) {

          case 'mis':
            filter =
              `SolicitanteId eq ${userId}`;
            break;

          case 'pendientes':
            filter =
              `Estado eq 'Enviado Aprobacion' and AprobadorId eq ${userId}`;
            break;

          case 'aprobadas':
            filter =
              `Estado eq 'Aprobado' and AprobadorId eq ${userId}`;
            break;

          case 'rechazadas':
            filter =
              `Estado eq 'Rechazado' and AprobadorId eq ${userId}`;
            break;

          case 'todas':
            filter = '';
            break;
        }

        let requerimientosUrl =
          `${context.pageContext.web.absoluteUrl}` +
          `/_api/web/lists/getbytitle('Requerimientos')/items` +
          `?$select=` +
          `Id,Title,Descripcion,Area,CategoriaId,AprobadorId,ValorPromedio,ValorTotal,Estado,` +
          `Recurrente,Created,EtapaActual,FechaAprobacion,FechaEnvioAprobacion,FechaDecision,ComentarioAprobador,` +
          `Aprobador/Id,Aprobador/Title,Aprobador/EMail,` +
          `DecisionPor/Id,DecisionPor/Title,DecisionPor/EMail,` +
          `Solicitante/Id,Solicitante/Title,Solicitante/EMail` +
          `&$expand=Solicitante,DecisionPor,Aprobador` +
          `&$orderby=Created desc` +
          `&$top=500`;

        if (filter) {
          requerimientosUrl +=
            `&$filter=${encodeURIComponent(filter)}`;
        }

        const response =
          await context.spHttpClient.get(
            requerimientosUrl,
            SPHttpClient.configurations.v1,
            {
              headers: {
                Accept:
                  'application/json;odata=nometadata'
              }
            }
          );

        if (!response.ok) {

          const errorText =
            await response.text();

          throw new Error(
            `Error cargando requerimientos. ` +
            `HTTP ${response.status}: ${errorText}`
          );
        }

        const data =
          await response.json();

        console.log(
          'REQUERIMIENTOS:',
          data.value
        );

        setRequerimientos(
          data.value as IRequerimientoItem[]
        );

      } catch (error) {

        console.error(
          'Error cargando requerimientos:',
          error
        );

        setMessage(
          getErrorMessage(error)
        );

      } finally {

        setLoadingRequerimientos(false);
      }
    };

  const getCotizacionesDetalle = async (
    requerimientoId: number
  ): Promise<ICotizacionDetalle[]> => {
    const cotizacionesUrl =
      `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Cotizaciones')/items` +
      `?$select=Id,Title,NumeroCotizacionProveedor,CodigoCotizacionInterno,ValorTotal,Seleccionada,` +
      `AttachmentFiles/FileName,AttachmentFiles/ServerRelativeUrl` +
      `&$expand=AttachmentFiles` +
      `&$filter=${encodeURIComponent(`RequerimientoId eq ${requerimientoId} and Activo eq 1`)}` +
      `&$orderby=Id asc`;

    const response = await context.spHttpClient.get(
      cotizacionesUrl,
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=nometadata' } }
    );

    if (!response.ok) {
      throw new Error(
        `Error cargando cotizaciones. HTTP ${response.status}: ${await response.text()}`
      );
    }

    const data = await response.json();

    return data.value.map((cotizacion: {
      Id: number;
      Title: string;
      NumeroCotizacionProveedor?: string;
      CodigoCotizacionInterno?: string;
      ValorTotal?: number;
      Seleccionada?: boolean;
      AttachmentFiles?: IAdjuntoCotizacion[];
    }): ICotizacionDetalle => ({
      Id: cotizacion.Id,
      Title: cotizacion.Title,
      NumeroCotizacionProveedor: cotizacion.NumeroCotizacionProveedor,
      CodigoCotizacionInterno: cotizacion.CodigoCotizacionInterno,
      ValorTotal: Number(cotizacion.ValorTotal || 0),
      Seleccionada: !!cotizacion.Seleccionada,
      Attachments: cotizacion.AttachmentFiles || []
    }));
  };

  const loadDetalleRequerimiento = async (
    requerimientoId: number
  ): Promise<void> => {
    try {
      setLoadingDetalle(true);
      setCotizacionesDetalle([]);
      setOrdenCompraDetalle(null);

      const [detalle, ordenCompra] = await Promise.all([
        getCotizacionesDetalle(requerimientoId),
        getOrdenCompraPorRequerimiento(requerimientoId)
      ]);

      setCotizacionesDetalle(detalle);
      setOrdenCompraDetalle(ordenCompra);

      const cotizacionYaSeleccionada = detalle.find(
        (cotizacion: ICotizacionDetalle) => cotizacion.Seleccionada
      );

      setCotizacionSeleccionadaId(cotizacionYaSeleccionada?.Id || null);
    } catch (error) {
      console.error('Error cargando detalle:', error);
      setMessage(getErrorMessage(error));
    } finally {
      setLoadingDetalle(false);
    }
  };


  const loadRequerimientoPorId = async (
    requerimientoId: number
  ): Promise<IRequerimientoItem> => {
    const requerimientoUrl =
      `${context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists/getbytitle('Requerimientos')/items(${requerimientoId})` +
      `?$select=` +
      `Id,Title,Descripcion,Area,CategoriaId,AprobadorId,ValorPromedio,ValorTotal,Estado,` +
      `Recurrente,Created,EtapaActual,FechaAprobacion,FechaEnvioAprobacion,FechaDecision,ComentarioAprobador,` +
      `Aprobador/Id,Aprobador/Title,Aprobador/EMail,` +
      `DecisionPor/Id,DecisionPor/Title,DecisionPor/EMail,` +
      `Solicitante/Id,Solicitante/Title,Solicitante/EMail` +
      `&$expand=Solicitante,DecisionPor,Aprobador`;

    const response = await context.spHttpClient.get(
      requerimientoUrl,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata'
        }
      }
    );

    if (response.status === 404) {
      throw new Error(`No existe el requerimiento #${requerimientoId}.`);
    }

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(
        `No se pudo abrir el requerimiento #${requerimientoId}. ` +
        `HTTP ${response.status}: ${errorText}`
      );
    }

    const data = await response.json();
    return data as IRequerimientoItem;
  };


  /*
   * ===========================
   * CARGA INICIAL
   * ===========================
   */

  useEffect(() => {

    const loadInitialData =
      async (): Promise<void> => {

        try {

          /*
           * Usuario actual de SharePoint
           */

          const currentUserUrl =
            `${context.pageContext.web.absoluteUrl}` +
            `/_api/web/currentuser` +
            `?$select=Id,Title,Email`;

          const currentUserResponse =
            await context.spHttpClient.get(
              currentUserUrl,
              SPHttpClient.configurations.v1,
              {
                headers: {
                  Accept:
                    'application/json;odata=nometadata'
                }
              }
            );

          if (!currentUserResponse.ok) {
            throw new Error(
              'No se pudo obtener el usuario actual de SharePoint.'
            );
          }

          const currentUser:
            ISharePointUser =
              await currentUserResponse.json();

          /*
           * Rol y área desde UsuariosRoles
           */

          const userConfig =
            await getCurrentUserConfiguration(
              currentUser.Id
            );

          const userRole = userConfig.role;

          setUser({
            id: currentUser.Id,
            displayName: currentUser.Title || 'Usuario',
            email: currentUser.Email || '',
            role: userRole,
            areaId: userConfig.areaId,
            area: userConfig.area,
            roleId: userConfig.roleId
          });

          /*
           * Categorías activas del área del usuario.
           * Admin puede visualizar todas las categorías.
           */

          await loadCategories(
            userConfig.areaId,
            userRole === 'Admin' ||
              userRole === 'Aprobador' ||
              userRole === 'ResponsableOC'
          );


          /*
           * Entrada directa desde correo:
           * CollabHome.aspx?reqId=123
           *
           * La navegación normal de la SPA no cambia la URL.
           * Este parámetro solo se utiliza al cargar inicialmente la página.
           */

          const rawReqId =
            new URLSearchParams(window.location.search).get('reqId');

          const deepLinkReqId =
            rawReqId ? Number(rawReqId) : undefined;

          let deepLinkError = '';

          if (
            rawReqId &&
            (
              !deepLinkReqId ||
              Math.floor(deepLinkReqId) !== deepLinkReqId ||
              deepLinkReqId <= 0
            )
          ) {
            deepLinkError =
              'El enlace recibido contiene un ID de requerimiento inválido.';
          }

          if (deepLinkReqId && !deepLinkError) {
            try {
              const requerimientoDirecto =
                await loadRequerimientoPorId(deepLinkReqId);

              const puedeAbrir =
                userRole === 'Admin' ||
                (
                  userRole === 'Aprobador' &&
                  requerimientoDirecto.AprobadorId === currentUser.Id &&
                  requerimientoDirecto.Estado !== 'Borrador'
                ) ||
                (
                  userRole === 'Solicitador' &&
                  requerimientoDirecto.Solicitante?.Id === currentUser.Id
                );

              if (!puedeAbrir) {
                throw new Error(
                  'No tienes permisos para visualizar este requerimiento.'
                );
              }

              let directView: AppView = 'mis';

              if (userRole === 'Admin') {
                directView = 'todas';
              } else if (userRole === 'Aprobador') {
                if (requerimientoDirecto.Estado === 'Aprobado') {
                  directView = 'aprobadas';
                } else if (requerimientoDirecto.Estado === 'Rechazado') {
                  directView = 'rechazadas';
                } else {
                  directView = 'pendientes';
                }
              }

              setActiveView(directView);

              await loadRequerimientos(
                directView,
                currentUser.Id
              );

              setSelectedRequerimiento(
                requerimientoDirecto
              );
              setEtapaConsulta(
                requerimientoDirecto.EtapaActual || 'Solicitud'
              );

              setComentarioAprobador('');
              setCotizacionSeleccionadaId(null);
              setArchivoPreview(null);

              await loadDetalleRequerimiento(
                requerimientoDirecto.Id
              );

              return;

            } catch (deepLinkException) {
              deepLinkError =
                getErrorMessage(deepLinkException);
            }
          }

          /*
           * Vista inicial normal según rol
           */

          if (
            userRole === 'Aprobador'
          ) {

            setActiveView(
              'pendientes'
            );

            await loadRequerimientos(
              'pendientes',
              currentUser.Id
            );

          } else if (
            userRole === 'Admin'
          ) {

            setActiveView(
              'todas'
            );

            await loadRequerimientos(
              'todas',
              currentUser.Id
            );

          } else if (
            userRole === 'ResponsableOC'
          ) {

            setActiveView(
              'ordenes'
            );

            await loadOrdenesCompra(
              currentUser.Id
            );

          } else {

            setActiveView(
              'nueva'
            );
          }

          if (deepLinkError) {
            setMessage(deepLinkError);
          }

        } catch (error) {

          console.error(
            'Error cargando información inicial:',
            error
          );

          setMessage(
            getErrorMessage(error)
          );

        } finally {

          setLoading(false);
        }
      };

    void loadInitialData();

  }, [context]);

  /*
   * ===========================
   * CAMBIAR DE VISTA
   * ===========================
   */

  const changeView =
    (view: AppView): void => {

      setActiveView(view);
      setSelectedRequerimiento(null);
      setSelectedOrdenCompra(null);
      setOrdenCompraDetalle(null);
      setEtapaConsulta(null);
      setMessage('');

      if (!user) {
        return;
      }

      if (view === 'ordenes') {
        void loadOrdenesCompra(user.id);
        return;
      }

      if (view !== 'nueva') {
        void loadRequerimientos(
          view,
          user.id
        );
      }
    };

  const formatCurrency = (value?: number | null): string =>
    Number(value || 0).toLocaleString('es-EC', {
      style: 'currency',
      currency: 'USD'
    });

  const formatDateTime = (value?: string): string => {
    if (!value) {
      return 'Fecha no registrada';
    }

    return new Date(value).toLocaleString('es-EC', {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit'
    });
  };


  // Formato provisional hasta que negocio confirme la estructura oficial.
  // Solo hay que cambiar esta función cuando se valide la numeración definitiva.
  const generarNumeroOCProvisional = (requerimientoId: number): string =>
    `OC-${('000000' + requerimientoId).slice(-6)}`;

  const getVistaPreviaArchivo = (archivo: IAdjuntoCotizacion): string =>
    new URL(archivo.ServerRelativeUrl, window.location.origin).href;

  const updateListItem = async (listTitle: string, itemId: number, body: Record<string, unknown>): Promise<void> => {
    const url = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listTitle}')/items(${itemId})`;
    const response = await context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
      body: JSON.stringify(body)
    });
    if (!response.ok) {
      throw new Error(`Error actualizando ${listTitle}. HTTP ${response.status}: ${await response.text()}`);
    }
  };

  const createListItem = async (
    listTitle: string,
    body: Record<string, unknown>
  ): Promise<number> => {
    const url =
      `${context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists/getbytitle('${listTitle}')/items`;

    const response = await context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata'
        },
        body: JSON.stringify(body)
      }
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(
        `Error creando elemento en ${listTitle}. HTTP ${response.status}: ${errorText}`
      );
    }

    const created = await response.json();
    const id = created.Id || created.ID;
    if (!id) {
      throw new Error(`Se creó el elemento en ${listTitle}, pero no se obtuvo su ID.`);
    }
    return Number(id);
  };


  const getResponsableOCPorArea = async (
    area?: string
  ): Promise<IResponsableOC> => {
    const nombreArea = String(area || '').trim();

    if (!nombreArea) {
      throw new Error(
        'El requerimiento no tiene un área configurada. No se puede asignar ResponsableOC.'
      );
    }

    const safeArea = nombreArea.replace(/'/g, "''");
    const filter = encodeURIComponent(
      `NombreArea eq '${safeArea}' and Activo eq 1`
    );

    const areaUrl =
      `${context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists/getbytitle('Areas')/items` +
      `?$select=Id,NombreArea,ResponsableOC/Id,ResponsableOC/Title,ResponsableOC/EMail` +
      `&$expand=ResponsableOC` +
      `&$filter=${filter}` +
      `&$top=2`;

    const response = await context.spHttpClient.get(
      areaUrl,
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=nometadata' } }
    );

    if (!response.ok) {
      throw new Error(
        `No se pudo consultar ResponsableOC para el área ${nombreArea}. ` +
        `HTTP ${response.status}: ${await response.text()}`
      );
    }

    const data = await response.json();
    const areas = data.value || [];

    if (areas.length === 0) {
      throw new Error(
        `No existe una configuración activa para el área "${nombreArea}".`
      );
    }

    if (areas.length > 1) {
      throw new Error(
        `Existe más de una configuración activa para el área "${nombreArea}".`
      );
    }

    const responsable = areas[0].ResponsableOC as IResponsableOC | undefined;

    if (!responsable?.Id) {
      throw new Error(
        `El área "${nombreArea}" no tiene ResponsableOC configurado.`
      );
    }

    return responsable;
  };


  const crearOrdenCompraSiNoExiste = async (
    requerimiento: IRequerimientoItem,
    cotizacionId: number,
    solicitadoPorId: number,
    responsableOCId: number
  ): Promise<void> => {
    const filter = encodeURIComponent(
      `RequerimientoId eq ${requerimiento.Id} and Activo eq 1`
    );

    const existingUrl =
      `${context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists/getbytitle('OrdenesCompra')/items` +
      `?$select=Id,ResponsableOCId,NumeroOC,EstadoOC,FechaGeneracion,TipoOrden` +
      `&$filter=${filter}&$top=1`;

    const existingResponse = await context.spHttpClient.get(
      existingUrl,
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=nometadata' } }
    );

    if (!existingResponse.ok) {
      throw new Error(
        `No se pudo verificar la orden de compra existente. HTTP ${existingResponse.status}: ${await existingResponse.text()}`
      );
    }

    const existingData = await existingResponse.json();
    const esGeneracionAutomatica = requerimiento.Recurrente;
    const ahora = new Date().toISOString();
    const numeroOCAutomatico = esGeneracionAutomatica
      ? generarNumeroOCProvisional(requerimiento.Id)
      : undefined;

    if (existingData.value?.length > 0) {
      const existingOrder = existingData.value[0] as {
        Id: number;
        ResponsableOCId?: number;
        NumeroOC?: string | null;
        EstadoOC?: string;
        FechaGeneracion?: string | null;
        TipoOrden?: string;
      };

      const cambios: Record<string, unknown> = {};

      // Conserva la asignación histórica si ya existía.
      if (!existingOrder.ResponsableOCId) {
        cambios.ResponsableOCId = responsableOCId;
      }

      // En un reintento de aprobación, completa la generación automática
      // si la OC de Contrato quedó incompleta.
      if (esGeneracionAutomatica) {
        if (!existingOrder.NumeroOC) {
          cambios.NumeroOC = numeroOCAutomatico;
        }

        if (existingOrder.EstadoOC !== 'Generada') {
          cambios.EstadoOC = 'Generada';
        }

        if (!existingOrder.FechaGeneracion) {
          cambios.FechaGeneracion = ahora;
        }
      }

      if (Object.keys(cambios).length > 0) {
        await updateListItem(
          'OrdenesCompra',
          existingOrder.Id,
          cambios
        );
      }

      return;
    }

    const ordenBody: Record<string, unknown> = {
      Title: `OC-REQ-${requerimiento.Id}`,
      RequerimientoId: requerimiento.Id,
      CotizacionId: cotizacionId,
      TipoOrden: esGeneracionAutomatica ? 'GenerarOC' : 'SolicitarOC',
      TipoDocumento: esGeneracionAutomatica ? 'Contrato' : 'OCV',
      EstadoOC: esGeneracionAutomatica ? 'Generada' : 'Pendiente',
      ResponsableOCId: responsableOCId,
      FechaSolicitud: ahora,
      SolicitadoPorId: solicitadoPorId,
      Activo: true
    };

    if (esGeneracionAutomatica) {
      ordenBody.NumeroOC = numeroOCAutomatico;
      ordenBody.FechaGeneracion = ahora;
    }

    await createListItem(
      'OrdenesCompra',
      ordenBody
    );
  };

  const decidirRequerimiento = async (aprobar: boolean): Promise<void> => {
    if (!selectedRequerimiento || !user) return;

    if (user.role !== 'Aprobador' && user.role !== 'Admin') {
      setMessage('No tienes permisos para tomar esta decisión.');
      return;
    }

    if (aprobar && !cotizacionSeleccionadaId) {
      setMessage('Seleccione una cotización antes de aprobar el requerimiento.');
      return;
    }

    if (!aprobar && !comentarioAprobador.trim()) {
      setMessage('Ingrese un comentario para rechazar el requerimiento.');
      return;
    }

    const requerimientoId = selectedRequerimiento.Id;

    try {
      setProcesandoDecision(true);
      setMessage('');

      // Siempre trabajar con el estado más reciente de SharePoint.
      const requerimientoActual =
        await loadRequerimientoPorId(requerimientoId);

      if (requerimientoActual.Estado !== 'Enviado Aprobacion') {
        throw new Error(
          `El requerimiento #${requerimientoId} ya no está pendiente de aprobación. ` +
          `Estado actual: ${requerimientoActual.Estado}.`
        );
      }

      // Un aprobador únicamente puede decidir requerimientos asignados a él.
      // Admin se mantiene como excepción administrativa.
      if (
        user.role === 'Aprobador' &&
        requerimientoActual.AprobadorId !== user.id
      ) {
        throw new Error(
          'Este requerimiento está asignado a otro aprobador.'
        );
      }

      const cotizacionesActuales =
        await getCotizacionesDetalle(requerimientoId);

      let valorCotizacionAprobada: number | undefined;

      if (aprobar) {
        if (!cotizacionSeleccionadaId) {
          throw new Error(
            'Seleccione una cotización antes de aprobar el requerimiento.'
          );
        }

        const cotizacionSeleccionada =
          cotizacionesActuales.find(
            (cotizacion: ICotizacionDetalle) =>
              cotizacion.Id === cotizacionSeleccionadaId
          );

        if (!cotizacionSeleccionada) {
          throw new Error(
            'La cotización seleccionada ya no se encuentra disponible.'
          );
        }

        if (
          !Number.isFinite(cotizacionSeleccionada.ValorTotal) ||
          cotizacionSeleccionada.ValorTotal <= 0
        ) {
          throw new Error(
            'La cotización seleccionada no tiene un valor válido.'
          );
        }

        if (cotizacionSeleccionada.Attachments.length === 0) {
          throw new Error(
            'La cotización seleccionada no tiene un PDF adjunto.'
          );
        }

        valorCotizacionAprobada = cotizacionSeleccionada.ValorTotal;

        // Validar la asignación operativa antes de modificar cotizaciones o crear la OC.
        const responsableOC = await getResponsableOCPorArea(
          requerimientoActual.Area
        );

        console.log('ASIGNACIÓN RESPONSABLE OC:', {
          requerimientoId,
          area: requerimientoActual.Area,
          responsableOCId: responsableOC.Id,
          responsableOC: responsableOC.Title,
          responsableOCEmail: responsableOC.EMail
        });

        const mensajeConfirmacion = requerimientoActual.Recurrente
          ? 'Al aprobar este requerimiento se generará automáticamente la orden de compra asociada a un Contrato con la cotización seleccionada. ¿Deseas continuar?'
          : 'Al aprobar este requerimiento se generará automáticamente la solicitud de orden de compra (OCV) con la cotización seleccionada. ¿Deseas continuar?';

        if (!window.confirm(mensajeConfirmacion)) {
          return;
        }

        // Garantiza que exista una sola cotización seleccionada
        // usando la información recién leída desde SharePoint.
        for (const cotizacion of cotizacionesActuales) {
          await updateListItem('Cotizaciones', cotizacion.Id, {
            Seleccionada:
              cotizacion.Id === cotizacionSeleccionadaId
          });
        }

        const solicitadoPorId =
          requerimientoActual.Solicitante?.Id;

        if (!solicitadoPorId) {
          throw new Error(
            'El requerimiento no tiene un solicitante válido.'
          );
        }

        await crearOrdenCompraSiNoExiste(
          requerimientoActual,
          cotizacionSeleccionadaId,
          solicitadoPorId,
          responsableOC.Id
        );
      } else {
        // Un requerimiento rechazado no debe conservar
        // una cotización marcada como seleccionada.
        for (const cotizacion of cotizacionesActuales) {
          if (cotizacion.Seleccionada) {
            await updateListItem(
              'Cotizaciones',
              cotizacion.Id,
              { Seleccionada: false }
            );
          }
        }
      }

      await updateListItem('Requerimientos', requerimientoId, {
        Estado: aprobar ? 'Aprobado' : 'Rechazado',
        EtapaActual: aprobar
          ? (
              requerimientoActual.Recurrente
                ? 'Facturacion'
                : 'SolicitarOC'
            )
          : 'Aprobacion',
        FechaDecision: new Date().toISOString(),
        FechaAprobacion: aprobar ? new Date().toISOString() : null,
        ComentarioAprobador: comentarioAprobador.trim(),
        DecisionPorId: user.id,
        ...(aprobar && valorCotizacionAprobada !== undefined
          ? { ValorTotal: valorCotizacionAprobada }
          : {})
      });

      setSelectedRequerimiento(null);
      setComentarioAprobador('');
      setCotizacionSeleccionadaId(null);

      await loadRequerimientos(activeView, user.id);

      setMessage(
        aprobar
          ? requerimientoActual.Recurrente
            ? `Requerimiento #${requerimientoId} aprobado. La orden de compra se generó automáticamente y el proceso pasó a Facturación.`
            : `Requerimiento #${requerimientoId} aprobado. La solicitud de orden de compra quedó pendiente de gestión en ERP.`
          : `Requerimiento #${requerimientoId} rechazado correctamente.`
      );
    } catch (error) {
      setMessage(getErrorMessage(error));
    } finally {
      setProcesandoDecision(false);
    }
  };

  const getEtapas = (item: IRequerimientoItem): string[] =>
    item.Recurrente
      ? ['Solicitud', 'Aprobacion', 'GenerarOC', 'Facturacion', 'Finalizado']
      : ['Solicitud', 'Aprobacion', 'SolicitarOC', 'Facturacion', 'Finalizado'];

  const getEtapaLabel = (etapa?: string): string => {
    const labels: Record<string, string> = {
      Solicitud: 'Solicitud',
      Aprobacion: 'Aprobación',
      GenerarOC: 'Orden de compra',
      SolicitarOC:
        user?.role === 'ResponsableOC'
          ? 'Ingresar orden de compra'
          : 'Orden de compra',
      Facturacion: 'Facturación',
      Finalizado: 'Finalizado'
    };
    return labels[etapa || ''] || etapa || 'Solicitud';
  };

  const getEstadoClass = (estado: string): string => {
    if (estado === 'Aprobado') return `${styles.statusBadge} ${styles.statusApproved}`;
    if (estado === 'Rechazado') return `${styles.statusBadge} ${styles.statusRejected}`;
    if (estado === 'Borrador') return `${styles.statusBadge} ${styles.statusDraft}`;
    return `${styles.statusBadge} ${styles.statusPending}`;
  };

  const getSelectedQuotation = (): ICotizacionDetalle | undefined =>
    cotizacionesDetalle.find(
      (cotizacion: ICotizacionDetalle) => cotizacion.Seleccionada
    );

  const getHistoryEvents = (
    item: IRequerimientoItem
  ): IHistoryEvent[] => {
    const events: IHistoryEvent[] = [
      {
        key: 'created',
        title: 'Solicitud creada',
        date: item.Created,
        actor: item.Solicitante?.Title || 'Solicitante',
        detail: 'Se creó el requerimiento.'
      }
    ];

    if (
      item.Estado !== 'Borrador' ||
      item.FechaEnvioAprobacion
    ) {
      events.push({
        key: 'submitted',
        title: 'Enviado a aprobación',
        date: item.FechaEnvioAprobacion,
        actor: item.Solicitante?.Title || 'Solicitante',
        detail: item.Aprobador?.Title
          ? `Asignado a ${item.Aprobador.Title}.`
          : 'Enviado al aprobador correspondiente.'
      });
    }

    if (
      item.Estado === 'Aprobado' ||
      item.Estado === 'Rechazado'
    ) {
      const selectedQuotation = getSelectedQuotation();
      const decisionDetails: string[] = [];

      if (
        item.Estado === 'Aprobado' &&
        selectedQuotation
      ) {
        decisionDetails.push(
          `${selectedQuotation.CodigoCotizacionInterno || selectedQuotation.Title}: ${formatCurrency(selectedQuotation.ValorTotal)}`
        );
      }

      if (item.ComentarioAprobador) {
        decisionDetails.push(item.ComentarioAprobador);
      }

      events.push({
        key: 'decision',
        title:
          item.Estado === 'Aprobado'
            ? 'Requerimiento aprobado'
            : 'Requerimiento rechazado',
        date:
          item.FechaDecision ||
          item.FechaAprobacion,
        actor:
          item.DecisionPor?.Title ||
          item.Aprobador?.Title ||
          'Aprobador',
        detail:
          decisionDetails.join(' · ') ||
          'Decisión registrada.'
      });
    }

    const orden =
      ordenCompraDetalle?.RequerimientoId === item.Id
        ? ordenCompraDetalle
        : null;

    if (orden) {
      if (orden.TipoOrden === 'GenerarOC') {
        const detallesOC: string[] = [];

        if (orden.NumeroOC) {
          detallesOC.push(`Número OC: ${orden.NumeroOC}`);
        }

        if (orden.TipoDocumento) {
          detallesOC.push(`Tipo: ${orden.TipoDocumento}`);
        }

        events.push({
          key: `oc-generated-${orden.Id}`,
          title: 'Orden de compra generada automáticamente',
          date: orden.FechaGeneracion || orden.FechaSolicitud,
          actor: 'Sistema',
          detail:
            detallesOC.join(' · ') ||
            'Orden de compra generada automáticamente.'
        });
      } else {
        const responsable =
          orden.ResponsableOC?.Title ||
          'Responsable de orden de compra';

        events.push({
          key: `oc-request-${orden.Id}`,
          title: 'Orden de compra asignada',
          date: orden.FechaSolicitud,
          actor: responsable,
          detail:
            `Pendiente de ingresar el número emitido por el ERP. ` +
            `Estado: ${orden.EstadoOC}.`
        });

        if (orden.EstadoOC === 'Generada' && orden.NumeroOC) {
          const detallesRegistro: string[] = [
            `Número OC: ${orden.NumeroOC}`
          ];

          if (orden.Observaciones) {
            detallesRegistro.push(orden.Observaciones);
          }

          events.push({
            key: `oc-registered-${orden.Id}`,
            title: 'Orden de compra registrada',
            date: orden.FechaGeneracion || orden.FechaSolicitud,
            actor: responsable,
            detail: detallesRegistro.join(' · ')
          });
        }
      }
    }

    if (orden?.AttachmentFiles?.length) {
      const factura = orden.AttachmentFiles[0];
      const responsable =
        orden.ResponsableOC?.Title ||
        'Responsable de orden de compra';

      events.push({
        key: `invoice-${orden.Id}`,
        title:
          item.EtapaActual === 'Finalizado'
            ? 'Factura cargada y requerimiento finalizado'
            : 'Factura PDF cargada',
        date: orden.Modified,
        actor: responsable,
        detail: `Archivo: ${factura.FileName}`
      });
    }

    return events;
  };

  const getStageInspection = (
    item: IRequerimientoItem,
    etapa: string
  ): IStageInspection => {
    const selectedQuotation = getSelectedQuotation();

    if (etapa === 'Solicitud') {
      return {
        title: 'Solicitud',
        status:
          item.Estado === 'Borrador'
            ? 'Borrador'
            : 'Completada',
        date: item.Created,
        actor: item.Solicitante?.Title || 'Solicitante',
        details: [
          `Categoría: ${getCategoriaNombre(item.CategoriaId)}`,
          `Tipo: ${item.Recurrente ? 'Recurrente' : 'No recurrente'}`,
          `Valor promedio: ${formatCurrency(item.ValorPromedio ?? item.ValorTotal ?? 0)}`
        ]
      };
    }

    if (etapa === 'Aprobacion') {
      const decisionMade =
        item.Estado === 'Aprobado' ||
        item.Estado === 'Rechazado';

      const details: string[] = [
        item.Aprobador?.Title
          ? `Aprobador asignado: ${item.Aprobador.Title}`
          : 'Aprobador asignado según categoría.'
      ];

      if (decisionMade && selectedQuotation && item.Estado === 'Aprobado') {
        details.push(
          `Cotización aprobada: ${selectedQuotation.CodigoCotizacionInterno || selectedQuotation.Title} · ${formatCurrency(selectedQuotation.ValorTotal)}`
        );
      }

      if (item.ComentarioAprobador) {
        details.push(`Comentario: ${item.ComentarioAprobador}`);
      }

      return {
        title: 'Aprobación',
        status: decisionMade
          ? item.Estado
          : item.Estado === 'Borrador'
            ? 'No iniciada'
            : 'Pendiente de decisión',
        date: decisionMade
          ? (item.FechaDecision || item.FechaAprobacion)
          : item.FechaEnvioAprobacion,
        actor: decisionMade
          ? (item.DecisionPor?.Title || item.Aprobador?.Title)
          : item.Aprobador?.Title,
        details
      };
    }

    if (etapa === 'GenerarOC' || etapa === 'SolicitarOC') {
      const orden =
        ordenCompraDetalle?.RequerimientoId === item.Id
          ? ordenCompraDetalle
          : null;

      if (!orden) {
        return {
          title: getEtapaLabel(etapa),
          status: 'Sin información',
          date:
            item.FechaDecision ||
            item.FechaAprobacion,
          actor: 'Proceso operativo',
          details: [
            'Todavía no existe una orden de compra asociada a este requerimiento.'
          ]
        };
      }

      if (orden.TipoOrden === 'GenerarOC') {
        const details: string[] = [
          'La orden de compra fue generada automáticamente al aprobar el requerimiento.'
        ];

        if (orden.NumeroOC) {
          details.push(`Número OC: ${orden.NumeroOC}`);
        }

        if (orden.TipoDocumento) {
          details.push(`Tipo: ${orden.TipoDocumento}`);
        }

        return {
          title: 'Orden de compra',
          status:
            orden.EstadoOC === 'Generada'
              ? 'Generada'
              : orden.EstadoOC,
          date:
            orden.FechaGeneracion ||
            orden.FechaSolicitud,
          actor: 'Sistema',
          details
        };
      }

      const responsable =
        orden.ResponsableOC?.Title ||
        'Responsable de orden de compra';

      const details: string[] = [];

      if (orden.EstadoOC === 'Generada' && orden.NumeroOC) {
        details.push(`Número OC: ${orden.NumeroOC}`);
        details.push(
          'La orden emitida por el ERP fue registrada correctamente.'
        );
      } else {
        details.push(
          'La orden debe gestionarse en el ERP y luego registrar aquí el número generado.'
        );
      }

      details.push(`Responsable: ${responsable}`);

      if (orden.Observaciones) {
        details.push(`Observaciones: ${orden.Observaciones}`);
      }

      return {
        title:
          user?.role === 'ResponsableOC'
            ? 'Ingresar orden de compra'
            : 'Orden de compra',
        status:
          orden.EstadoOC === 'Generada'
            ? 'OC registrada'
            : orden.EstadoOC === 'En proceso'
              ? 'En gestión'
              : 'Pendiente de ingresar OC',
        date:
          orden.EstadoOC === 'Generada'
            ? (orden.FechaGeneracion || orden.FechaSolicitud)
            : orden.FechaSolicitud,
        actor: responsable,
        details
      };
    }

    if (etapa === 'Facturacion') {
      const facturaAdjunta = ordenCompraDetalle?.AttachmentFiles?.[0];

      const details: string[] = facturaAdjunta
        ? [
            `Factura PDF: ${facturaAdjunta.FileName}`,
            item.EtapaActual === 'Finalizado'
              ? 'La factura fue cargada y el requerimiento quedó finalizado.'
              : 'La factura ya está cargada. Falta completar la finalización del requerimiento.'
          ]
        : [
            'Pendiente de cargar el PDF de la factura asociada a la orden de compra.'
          ];

      return {
        title: 'Facturación',
        status:
          item.EtapaActual === 'Finalizado'
            ? 'Completada'
            : item.EtapaActual === 'Facturacion'
              ? (facturaAdjunta ? 'Factura cargada' : 'Pendiente de factura')
              : 'No iniciada',
        date:
          facturaAdjunta
            ? ordenCompraDetalle?.Modified
            : undefined,
        actor:
          ordenCompraDetalle?.ResponsableOC?.Title ||
          'Responsable de orden de compra',
        details
      };
    }

    return {
      title: 'Finalizado',
      status:
        item.EtapaActual === 'Finalizado'
          ? 'Completado'
          : 'No iniciado',
      date:
        ordenCompraDetalle?.AttachmentFiles?.length
          ? ordenCompraDetalle.Modified
          : undefined,
      actor: 'Proceso',
      details: [
        ordenCompraDetalle?.AttachmentFiles?.length
          ? 'El requerimiento fue cerrado después de cargar el PDF de la factura.'
          : 'El requerimiento queda cerrado cuando termina la facturación.'
      ]
    };
  };

  const filteredRequerimientos = requerimientos.filter((item) => {
    const term = searchTerm.trim().toLowerCase();
    if (!term) return true;
    return (
      String(item.Id).indexOf(term) >= 0 ||
      item.Descripcion.toLowerCase().indexOf(term) >= 0 ||
      getCategoriaNombre(item.CategoriaId).toLowerCase().indexOf(term) >= 0 ||
      item.Estado.toLowerCase().indexOf(term) >= 0
    );
  });


  const filteredOrdenesCompra = ordenesCompra.filter((orden: IOrdenCompra) => {
    const term = searchTerm.trim().toLowerCase();

    if (!term) {
      return true;
    }

    return (
      String(orden.Id).indexOf(term) >= 0 ||
      String(orden.RequerimientoId || '').indexOf(term) >= 0 ||
      String(orden.Title || '').toLowerCase().indexOf(term) >= 0 ||
      String(orden.TipoDocumento || '').toLowerCase().indexOf(term) >= 0 ||
      String(orden.TipoOrden || '').toLowerCase().indexOf(term) >= 0 ||
      String(orden.EstadoOC || '').toLowerCase().indexOf(term) >= 0 ||
      String(orden.NumeroOC || '').toLowerCase().indexOf(term) >= 0
    );
  });


  const abrirOrdenCompra = async (
    orden: IOrdenCompra
  ): Promise<void> => {
    try {
      setSelectedOrdenCompra(orden);
      setOrdenCompraDetalle(orden);
      setNumeroOCInput(orden.NumeroOC || '');
      setObservacionesOCInput(orden.Observaciones || '');
      setFacturaPdf(null);
      setMessage('');

      if (!orden.RequerimientoId) {
        throw new Error(
          'La orden de compra no tiene un requerimiento asociado.'
        );
      }

      const requerimiento =
        await loadRequerimientoPorId(orden.RequerimientoId);

      setSelectedRequerimiento(requerimiento);
      setEtapaConsulta(
        requerimiento.EtapaActual || 'Solicitud'
      );

      await loadDetalleRequerimiento(requerimiento.Id);

      const ordenActualizada =
        await getOrdenCompraPorRequerimiento(requerimiento.Id);

      if (ordenActualizada) {
        setSelectedOrdenCompra(ordenActualizada);
        setOrdenCompraDetalle(ordenActualizada);
        setNumeroOCInput(ordenActualizada.NumeroOC || '');
        setObservacionesOCInput(
          ordenActualizada.Observaciones || ''
        );
      }
    } catch (error) {
      setMessage(getErrorMessage(error));
    }
  };

  const cerrarOrdenCompra = (): void => {
    setSelectedOrdenCompra(null);
    setSelectedRequerimiento(null);
    setOrdenCompraDetalle(null);
    setCotizacionesDetalle([]);
    setEtapaConsulta(null);
    setNumeroOCInput('');
    setObservacionesOCInput('');
    setFacturaPdf(null);
    setMessage('');
  };

  const registrarNumeroOrdenCompra = async (): Promise<void> => {
    if (!user || user.role !== 'ResponsableOC') {
      setMessage('No tienes permisos para gestionar órdenes de compra.');
      return;
    }

    if (!selectedOrdenCompra) {
      return;
    }

    if (
      selectedOrdenCompra.ResponsableOCId &&
      selectedOrdenCompra.ResponsableOCId !== user.id
    ) {
      setMessage('Esta orden de compra está asignada a otro responsable.');
      return;
    }

    if (selectedOrdenCompra.TipoOrden !== 'SolicitarOC') {
      setMessage(
        'Esta orden se genera automáticamente y no requiere registrar un número desde ERP.'
      );
      return;
    }

    if (selectedOrdenCompra.EstadoOC === 'Generada') {
      setMessage('Esta orden de compra ya fue registrada como generada.');
      return;
    }

    const numeroOC = numeroOCInput.trim();

    if (!numeroOC) {
      setMessage('Ingrese el número de orden de compra generado por el ERP.');
      return;
    }

    try {
      setProcesandoOrdenCompra(true);
      setMessage('');

      const ahora = new Date().toISOString();

      await updateListItem(
        'OrdenesCompra',
        selectedOrdenCompra.Id,
        {
          NumeroOC: numeroOC,
          EstadoOC: 'Generada',
          FechaGeneracion: ahora,
          Observaciones: observacionesOCInput.trim()
        }
      );

      if (selectedOrdenCompra.RequerimientoId) {
        await updateListItem(
          'Requerimientos',
          selectedOrdenCompra.RequerimientoId,
          {
            EtapaActual: 'Facturacion'
          }
        );
      }

      await loadOrdenesCompra(user.id);

      const ordenActualizada =
        selectedOrdenCompra.RequerimientoId
          ? await getOrdenCompraPorRequerimiento(
              selectedOrdenCompra.RequerimientoId
            )
          : null;

      if (ordenActualizada) {
        setSelectedOrdenCompra(ordenActualizada);
        setOrdenCompraDetalle(ordenActualizada);
        setNumeroOCInput(ordenActualizada.NumeroOC || '');
        setObservacionesOCInput(
          ordenActualizada.Observaciones || ''
        );
      }

      if (selectedOrdenCompra.RequerimientoId) {
        const requerimientoActualizado =
          await loadRequerimientoPorId(
            selectedOrdenCompra.RequerimientoId
          );

        setSelectedRequerimiento(requerimientoActualizado);
        setEtapaConsulta('Facturacion');
      }

      setMessage(
        `Orden ${numeroOC} registrada correctamente. El requerimiento pasó a Facturación.`
      );
    } catch (error) {
      setMessage(getErrorMessage(error));
    } finally {
      setProcesandoOrdenCompra(false);
    }
  };


  const adjuntarArchivo = async (
    listTitle: string,
    itemId: number,
    archivoPdf: File
  ): Promise<void> => {
    const safeFileName = archivoPdf.name.replace(/'/g, "''");
    const url =
      `${context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists/getbytitle('${listTitle}')/items(${itemId})` +
      `/AttachmentFiles/add(FileName=@fileName)` +
      `?@fileName='${encodeURIComponent(safeFileName)}'`;

    const response = await context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: { Accept: 'application/json;odata=nometadata' },
        body: archivoPdf
      }
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(
        `No se pudo adjuntar ${archivoPdf.name} en ${listTitle}. HTTP ${response.status}: ${errorText}`
      );
    }
  };


  const registrarFacturaYFinalizar = async (): Promise<void> => {
    if (!user || user.role !== 'ResponsableOC') {
      setMessage('No tienes permisos para registrar la factura.');
      return;
    }

    if (!selectedOrdenCompra || !selectedRequerimiento) {
      setMessage('No se pudo identificar la orden de compra o el requerimiento.');
      return;
    }

    if (
      selectedOrdenCompra.ResponsableOCId &&
      selectedOrdenCompra.ResponsableOCId !== user.id
    ) {
      setMessage('Esta orden de compra está asignada a otro responsable.');
      return;
    }

    if (selectedOrdenCompra.EstadoOC !== 'Generada') {
      setMessage('La orden de compra debe estar generada antes de registrar la factura.');
      return;
    }

    if (selectedRequerimiento.EtapaActual !== 'Facturacion') {
      setMessage('El requerimiento no se encuentra en la etapa de Facturación.');
      return;
    }

    const facturaExistente = selectedOrdenCompra.AttachmentFiles?.[0];

    if (!facturaExistente && !facturaPdf) {
      setMessage('Adjunte el PDF de la factura antes de finalizar.');
      return;
    }

    if (facturaPdf) {
      const esPdf =
        facturaPdf.type === 'application/pdf' ||
        /\.pdf$/i.test(facturaPdf.name);

      if (!esPdf) {
        setMessage('La factura debe ser un archivo PDF.');
        return;
      }
    }

    try {
      setProcesandoFactura(true);
      setMessage('');

      // Si el archivo ya existe por un intento anterior, no lo vuelve a adjuntar.
      if (!facturaExistente && facturaPdf) {
        await adjuntarArchivo(
          'OrdenesCompra',
          selectedOrdenCompra.Id,
          facturaPdf
        );
      }

      if (!selectedOrdenCompra.RequerimientoId) {
        throw new Error(
          'La orden de compra no tiene un requerimiento asociado.'
        );
      }

      await updateListItem(
        'Requerimientos',
        selectedOrdenCompra.RequerimientoId,
        {
          EtapaActual: 'Finalizado'
        }
      );

      const [ordenActualizada, requerimientoActualizado] = await Promise.all([
        getOrdenCompraPorRequerimiento(selectedOrdenCompra.RequerimientoId),
        loadRequerimientoPorId(selectedOrdenCompra.RequerimientoId)
      ]);

      if (ordenActualizada) {
        setSelectedOrdenCompra(ordenActualizada);
        setOrdenCompraDetalle(ordenActualizada);
      }

      setSelectedRequerimiento(requerimientoActualizado);
      setEtapaConsulta('Finalizado');
      setFacturaPdf(null);

      await loadOrdenesCompra(user.id);

      setMessage(
        'Factura registrada correctamente. El requerimiento quedó finalizado.'
      );
    } catch (error) {
      setMessage(getErrorMessage(error));
    } finally {
      setProcesandoFactura(false);
    }
  };


  const abrirRequerimiento =
  (requerimiento: IRequerimientoItem): void => {

    setSelectedRequerimiento(requerimiento);
    setSelectedOrdenCompra(null);
    setOrdenCompraDetalle(null);
    setEtapaConsulta(requerimiento.EtapaActual || 'Solicitud');
    setComentarioAprobador('');
    setCotizacionSeleccionadaId(null);
    setArchivoPreview(null);
    setMessage('');

    void loadDetalleRequerimiento(
      requerimiento.Id
    );
  };

  const continuarBorrador = async (
    requerimiento: IRequerimientoItem
  ): Promise<void> => {
    if (!user || requerimiento.Estado !== 'Borrador') {
      return;
    }

    const puedeEditar =
      user.role === 'Admin' ||
      (
        user.role === 'Solicitador' &&
        requerimiento.Solicitante?.Id === user.id
      );

    if (!puedeEditar) {
      setMessage('No tienes permisos para editar este borrador.');
      return;
    }

    try {
      setLoadingDetalle(true);
      setMessage('');

      const detalle = await getCotizacionesDetalle(requerimiento.Id);

      const cotizacionesBorrador: ICotizacionForm[] =
        detalle.length > 0
          ? detalle.map((cotizacion: ICotizacionDetalle) => ({
              key: `draft-${cotizacion.Id}`,
              id: cotizacion.Id,
              valorTotal:
                cotizacion.ValorTotal > 0
                  ? String(cotizacion.ValorTotal)
                  : '',
              archivo: null,
              adjuntosExistentes: cotizacion.Attachments
            }))
          : [createEmptyCotizacion()];

      setEditingDraftId(requerimiento.Id);
      setOriginalDraftCotizacionIds(
        detalle.map((cotizacion: ICotizacionDetalle) => cotizacion.Id)
      );
      setDescripcion(requerimiento.Descripcion || '');
      setCategoriaId(
        requerimiento.CategoriaId
          ? String(requerimiento.CategoriaId)
          : ''
      );
      setRecurrente(requerimiento.Recurrente);
      setCotizaciones(cotizacionesBorrador);
      setCotizacionesDetalle(detalle);
      setSelectedRequerimiento(null);
      setArchivoPreview(null);
      setCotizacionSeleccionadaId(null);
      setActiveView('nueva');
    } catch (error) {
      setMessage(getErrorMessage(error));
    } finally {
      setLoadingDetalle(false);
    }
  };

  const cancelarEdicionBorrador = (): void => {
    setEditingDraftId(null);
    setOriginalDraftCotizacionIds([]);
    setDescripcion('');
    setCategoriaId('');
    setRecurrente(false);
    setCotizaciones([createEmptyCotizacion()]);
    setMessage('');

    if (user) {
      const destino: AppView =
        user.role === 'Admin'
          ? 'todas'
          : 'mis';

      setActiveView(destino);
      void loadRequerimientos(destino, user.id);
    }
  };

  /*
   * ===========================
   * GUARDAR / ENVIAR REQUERIMIENTO
   * ===========================
   */

  const agregarCotizacion = (): void => {
    if (cotizaciones.length >= 3) {
      setMessage('Solo se permiten hasta 3 cotizaciones por requerimiento.');
      return;
    }
    setMessage('');
    setCotizaciones([...cotizaciones, createEmptyCotizacion()]);
  };

  const eliminarCotizacion = (cotizacionIndex: number): void => {
    if (cotizaciones.length === 1) {
      setMessage('Debe existir al menos una cotización en el formulario.');
      return;
    }
    setCotizaciones(cotizaciones.filter((_, index) => index !== cotizacionIndex));
  };

  const actualizarCotizacion = (cotizacionIndex: number, changes: Partial<ICotizacionForm>): void => {
    setCotizaciones(cotizaciones.map((cotizacion, index) =>
      index === cotizacionIndex ? { ...cotizacion, ...changes } : cotizacion
    ));
  };

  const saveRequerimiento = async (enviar: boolean): Promise<void> => {
    try {
      setSaving(true);
      setMessage('');

      if (!user) {
        throw new Error('No se pudo obtener el usuario actual.');
      }

      if (user.role !== 'Solicitador' && user.role !== 'Admin') {
        throw new Error('No tienes permisos para crear requerimientos.');
      }

      if (!descripcion.trim()) {
        throw new Error('Ingrese una descripción.');
      }

      if (!categoriaId) {
        throw new Error('Seleccione una categoría.');
      }

      const categoriaSeleccionada = categories.find(
        (categoria: ICategory) =>
          categoria.Id === Number(categoriaId)
      );

      if (!categoriaSeleccionada) {
        throw new Error(
          'No se pudo obtener la categoría seleccionada.'
        );
      }

      const aprobadorId = MODO_PRUEBAS
        ? categoriaSeleccionada.AprobadorPruebaId
        : categoriaSeleccionada.AprobadorCategoriaId;

      console.log('ENRUTAMIENTO APROBADOR:', {
        modo: MODO_PRUEBAS ? 'PRUEBAS' : 'PRODUCCION',
        categoriaId: categoriaSeleccionada.Id,
        categoria: categoriaSeleccionada.NombreCategoria,
        aprobadorCategoriaId: categoriaSeleccionada.AprobadorCategoriaId,
        aprobadorPruebaId: categoriaSeleccionada.AprobadorPruebaId,
        aprobadorElegidoId: aprobadorId
      });

      if (enviar && !aprobadorId) {
        throw new Error(
          MODO_PRUEBAS
            ? 'La categoría seleccionada no tiene un aprobador de prueba configurado.'
            : 'La categoría seleccionada no tiene un aprobador configurado.'
        );
      }

      if (cotizaciones.length < 1 || cotizaciones.length > 3) {
        throw new Error(
          'El requerimiento debe tener entre 1 y 3 cotizaciones.'
        );
      }

      if (enviar) {
        cotizaciones.forEach(
          (
            cotizacion: ICotizacionForm,
            cotizacionIndex: number
          ) => {
            const totalCotizacion =
              Number(cotizacion.valorTotal);

            if (
              !cotizacion.valorTotal ||
              isNaN(totalCotizacion) ||
              totalCotizacion <= 0
            ) {
              throw new Error(
                `Ingrese un valor total válido en la cotización ${cotizacionIndex + 1}.`
              );
            }

            const tienePdf =
              !!cotizacion.archivo ||
              cotizacion.adjuntosExistentes.length > 0;

            if (!tienePdf) {
              throw new Error(
                `Adjunte el PDF de la cotización ${cotizacionIndex + 1}.`
              );
            }
          }
        );
      }

      const valoresCotizaciones = cotizaciones
        .map((cotizacion: ICotizacionForm) =>
          Number(cotizacion.valorTotal)
        )
        .filter(
          (valor: number) =>
            !isNaN(valor) && valor > 0
        );

      const valorPromedio =
        valoresCotizaciones.length > 0
          ? valoresCotizaciones.reduce(
              (total: number, valor: number) =>
                total + valor,
              0
            ) / valoresCotizaciones.length
          : 0;

      const esEdicionBorrador =
        editingDraftId !== null;

      let itemId: number;

      if (esEdicionBorrador && editingDraftId) {
        itemId = editingDraftId;

        await updateListItem(
          'Requerimientos',
          itemId,
          {
            Title:
              descripcion.trim().substring(0, 255),
            Descripcion: descripcion.trim(),
            Area: user.area || '',
            CategoriaId: Number(categoriaId),
            Recurrente: recurrente,
            ValorPromedio: valorPromedio,
            ValorTotal: null,
            EtapaActual: 'Solicitud',
            Estado: 'Borrador',
            AprobadorId: null
          }
        );

        const idsActuales = cotizaciones
          .map(
            (cotizacion: ICotizacionForm) =>
              cotizacion.id
          )
          .filter(
            (id: number | undefined): id is number =>
              typeof id === 'number'
          );

        const idsEliminados =
          originalDraftCotizacionIds.filter(
            (id: number) =>
              idsActuales.indexOf(id) < 0
          );

        for (const cotizacionId of idsEliminados) {
          await updateListItem(
            'Cotizaciones',
            cotizacionId,
            { Activo: false }
          );
        }
      } else {
        itemId = await createListItem(
          'Requerimientos',
          {
            Title:
              descripcion.trim().substring(0, 255),
            Descripcion: descripcion.trim(),
            SolicitanteId: user.id,
            Area: user.area || '',
            CategoriaId: Number(categoriaId),
            AprobadorId: null,
            Recurrente: recurrente,
            ValorPromedio: valorPromedio,
            ValorTotal: null,
            EtapaActual: 'Solicitud',
            Estado: 'Borrador'
          }
        );
      }

      for (
        let c = 0;
        c < cotizaciones.length;
        c += 1
      ) {
        const cotizacion = cotizaciones[c];

        const tieneDatos =
          !!cotizacion.archivo ||
          cotizacion.adjuntosExistentes.length > 0 ||
          !!cotizacion.valorTotal;

        if (!enviar && !tieneDatos) {
          continue;
        }

        const totalCotizacion =
          Number(cotizacion.valorTotal) || 0;

        let cotizacionId =
          cotizacion.id;

        if (cotizacionId) {
          await updateListItem(
            'Cotizaciones',
            cotizacionId,
            {
              ValorTotal: totalCotizacion,
              Seleccionada: false,
              Activo: true
            }
          );
        } else {
          const cotizacionBody:
            Record<string, unknown> = {
              Title:
                `REQ-${itemId}-COT-${c + 1}`,
              RequerimientoId: itemId,
              CodigoCotizacionInterno:
                `COT-REQ-${('000000' + itemId).slice(-6)}-${('00' + (c + 1)).slice(-2)}`,
              ValorTotal: totalCotizacion,
              Seleccionada: false,
              Activo: true
            };

          cotizacionId =
            await createListItem(
              'Cotizaciones',
              cotizacionBody
            );
        }

        if (cotizacion.archivo) {
          const isPdf =
            cotizacion.archivo.type ===
              'application/pdf' ||
            /\.pdf$/i.test(
              cotizacion.archivo.name
            );

          if (!isPdf) {
            throw new Error(
              `El archivo de la cotización ${c + 1} debe ser PDF.`
            );
          }

          await adjuntarArchivo(
            'Cotizaciones',
            cotizacionId,
            cotizacion.archivo
          );
        }
      }

      if (enviar && aprobadorId) {
        await updateListItem(
          'Requerimientos',
          itemId,
          {
            Area: user.area || '',
            CategoriaId: Number(categoriaId),
            AprobadorId: aprobadorId,
            ValorPromedio: valorPromedio,
            ValorTotal: null,
            FechaEnvioAprobacion: new Date().toISOString(),
            EtapaActual: 'Aprobacion',
            Estado: 'Enviado Aprobacion'
          }
        );
      }

      setDescripcion('');
      setCategoriaId('');
      setRecurrente(false);
      setCotizaciones([createEmptyCotizacion()]);
      setEditingDraftId(null);
      setOriginalDraftCotizacionIds([]);

      if (user.role === 'Solicitador') {
        setActiveView('mis');
        await loadRequerimientos(
          'mis',
          user.id
        );
      } else {
        setActiveView('todas');
        await loadRequerimientos(
          'todas',
          user.id
        );
      }

      setMessage(
        enviar
          ? `Requerimiento #${itemId} enviado a aprobación correctamente.`
          : esEdicionBorrador
            ? `Borrador #${itemId} actualizado correctamente.`
            : `Requerimiento #${itemId} guardado como borrador.`
      );
    } catch (error) {
      setMessage(getErrorMessage(error));
    } finally {
      setSaving(false);
    }
  };

  /*
   * ===========================
   * RENDER
   * ===========================
   */

  if (loading) {
    return (
      <div className={styles.requerimientos} aria-busy="true" aria-label="Cargando información">
        <div className={styles.loadingSkeleton}>
          <div className={styles.skeletonTop} />
          <div className={styles.skeletonTitle} />
          <div className={styles.skeletonTabs} />
          <div className={styles.skeletonPanel} />
        </div>
      </div>
    );
  }

  return (

    <div className={styles.requerimientos}>

      {/* CABECERA COMPACTA */}

      <div className={styles.topBar}>
        <div className={styles.brand}>
          <img
            src={inovaLogo}
            alt="Inova Solutions"
            className={styles.logo}
          />
          <div className={styles.brandText}>
            <span className={styles.eyebrow}>OPERACIONES</span>
            <h1 className={styles.title}>Gestión de Requerimientos</h1>
          </div>
        </div>

        <div className={styles.headerUser}>
          {user && (
            <>
              <div className={styles.headerUserText}>
                <strong>{user.displayName}</strong>
                <span>{user.area ? `${user.area} · ${user.role}` : user.role}</span>
              </div>

              <img
                className={styles.avatarSmall}
                src={
                  `${context.pageContext.web.absoluteUrl}` +
                  `/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(user.email)}`
                }
                alt={user.displayName}
              />
            </>
          )}
        </div>
      </div>

      {/* SIN ACCESO */}

      {user?.role === 'SinRol' && (

        <div className={styles.emptyState}>
          No tienes un rol autorizado
          para utilizar esta aplicación.
        </div>
      )}

      {/* NAVEGACIÓN SOLICITADOR */}

      {user?.role === 'Solicitador' && (

        <div className={styles.tabs}>

          <button
            type="button"
            className={
              getTabClass('nueva')
            }
            onClick={() =>
              changeView('nueva')
            }
          >
            Nueva solicitud
          </button>

          <button
            type="button"
            className={
              getTabClass('mis')
            }
            onClick={() =>
              changeView('mis')
            }
          >
            Mis solicitudes
          </button>

        </div>
      )}

      {/* NAVEGACIÓN APROBADOR */}

      {user?.role === 'Aprobador' && (

        <div className={styles.tabs}>

          <button
            type="button"
            className={
              getTabClass('pendientes')
            }
            onClick={() =>
              changeView('pendientes')
            }
          >
            Pendientes
          </button>

          <button
            type="button"
            className={
              getTabClass('aprobadas')
            }
            onClick={() =>
              changeView('aprobadas')
            }
          >
            Aprobadas
          </button>

          <button
            type="button"
            className={
              getTabClass('rechazadas')
            }
            onClick={() =>
              changeView('rechazadas')
            }
          >
            Rechazadas
          </button>

        </div>
      )}

      {/* NAVEGACIÓN RESPONSABLE OC */}

      {user?.role === 'ResponsableOC' && (
        <div className={styles.tabs}>
          <button
            type="button"
            className={getTabClass('ordenes')}
            onClick={() => changeView('ordenes')}
          >
            Órdenes de compra
          </button>
        </div>
      )}

      {/* NAVEGACIÓN ADMIN */}

      {user?.role === 'Admin' && (

        <div className={styles.tabs}>

          <button
            type="button"
            className={
              getTabClass('todas')
            }
            onClick={() =>
              changeView('todas')
            }
          >
            Todas las solicitudes
          </button>

          <button
            type="button"
            className={
              getTabClass('nueva')
            }
            onClick={() =>
              changeView('nueva')
            }
          >
            Nueva solicitud
          </button>

        </div>
      )}

      {/* FORMULARIO NUEVA SOLICITUD */}

      {(
        user?.role === 'Solicitador' ||
        user?.role === 'Admin'
      ) && activeView === 'nueva' && (
        <div className={styles.formWorkspace}>
          <div className={styles.compactFormIntro}>
            <h2 className={styles.pageTitle}>
              {editingDraftId
                ? `Editar borrador #${editingDraftId}`
                : 'Nueva solicitud'}
            </h2>

            {editingDraftId && (
              <button
                type="button"
                className={styles.backButton}
                disabled={saving}
                onClick={cancelarEdicionBorrador}
              >
                Cancelar edición
              </button>
            )}
          </div>

          <section className={`${styles.formSection} ${styles.compactRequestSection}`}>
            <div className={styles.compactRequestGrid}>
              <div className={styles.formGroup}>
                <label className={styles.label} htmlFor="requerimiento-categoria">Categoría</label>
                <select
                  id="requerimiento-categoria"
                  className={styles.select}
                  value={categoriaId}
                  onChange={(event: React.ChangeEvent<HTMLSelectElement>) =>
                    setCategoriaId(event.target.value)
                  }
                >
                  <option value="">Seleccione una categoría</option>
                  {categories.map((category: ICategory) => (
                    <option key={category.Id} value={category.Id}>
                      {category.NombreCategoria}
                    </option>
                  ))}
                </select>
              </div>

              <div className={styles.formGroup}>
                <label className={styles.label}>Tipo</label>
                <div className={styles.compactProcessOptions}>
                  <button
                    type="button"
                    className={`${styles.compactProcessOption} ${!recurrente ? styles.compactProcessOptionActive : ''}`}
                    aria-pressed={!recurrente}
                    onClick={() => setRecurrente(false)}
                  >
                    <span className={styles.toggleDot} />
                    <span>
                      <strong>No recurrente</strong>
                      <small>OCV</small>
                    </span>
                  </button>

                  <button
                    type="button"
                    className={`${styles.compactProcessOption} ${recurrente ? styles.compactProcessOptionActive : ''}`}
                    aria-pressed={recurrente}
                    onClick={() => setRecurrente(true)}
                  >
                    <span className={styles.toggleDot} />
                    <span>
                      <strong>Recurrente</strong>
                      <small>Contrato</small>
                    </span>
                  </button>
                </div>
              </div>

              <div className={`${styles.formGroup} ${styles.compactDescription}`}>
                <label className={styles.label} htmlFor="requerimiento-descripcion">Descripción</label>
                <textarea
                  id="requerimiento-descripcion"
                  className={styles.textarea}
                  value={descripcion}
                  placeholder="Describe la necesidad..."
                  onChange={(event: React.ChangeEvent<HTMLTextAreaElement>) =>
                    setDescripcion(event.target.value)
                  }
                  rows={3}
                />
              </div>
            </div>
          </section>

          <section className={styles.formSection}>
            <div className={`${styles.sectionHeaderRow} ${styles.compactQuotationHeader}`}>
              <h3>Cotizaciones</h3>
              <span className={styles.counterBadge}>{cotizaciones.length} / 3</span>
            </div>

            <div className={styles.quotationList}>
              {cotizaciones.map((cotizacion, cotizacionIndex) => (
                <div key={cotizacion.key} className={styles.quotationCard}>
                  <div className={styles.quotationHeader}>
                    <span className={styles.quotationLabel}>
                      COTIZACIÓN {String(cotizacionIndex + 1)}
                    </span>

                    {cotizaciones.length > 1 && (
                      <button
                        type="button"
                        className={styles.quotationDeleteButton}
                        onClick={() => eliminarCotizacion(cotizacionIndex)}
                        aria-label={`Eliminar cotización ${cotizacionIndex + 1}`}
                        title={`Eliminar cotización ${cotizacionIndex + 1}`}
                      >
                        <span aria-hidden="true">🗑</span>
                      </button>
                    )}
                  </div>

                  <div className={styles.quotationBody}>
                    <div className={styles.formGroup}>
                      <label className={styles.label}>PDF</label>

                      {cotizacion.archivo ? (
                        <div className={styles.pdfFileChip}>
                          <span className={styles.pdfFileIcon} aria-hidden="true">PDF</span>
                          <span
                            className={styles.pdfFileName}
                            title={cotizacion.archivo.name}
                          >
                            {cotizacion.archivo.name}
                          </span>
                          <button
                            type="button"
                            className={styles.pdfFileRemove}
                            onClick={() =>
                              actualizarCotizacion(cotizacionIndex, { archivo: null })
                            }
                            aria-label={`Quitar PDF de cotización ${cotizacionIndex + 1}`}
                            title="Quitar archivo"
                          >
                            ×
                          </button>
                        </div>
                      ) : cotizacion.adjuntosExistentes.length > 0 ? (
                        <div className={styles.pdfFileChip}>
                          <span className={styles.pdfFileIcon} aria-hidden="true">PDF</span>
                          <span
                            className={styles.pdfFileName}
                            title={cotizacion.adjuntosExistentes[0].FileName}
                          >
                            {cotizacion.adjuntosExistentes[0].FileName}
                          </span>
                          <span title="Archivo ya guardado en SharePoint">
                            ✓
                          </span>
                        </div>
                      ) : (
                        <label className={styles.fileUpload}>
                          <input
                            type="file"
                            accept="application/pdf,.pdf"
                            aria-label={`Adjuntar PDF de cotización ${cotizacionIndex + 1}`}
                            onChange={(event: React.ChangeEvent<HTMLInputElement>) => {
                              const selectedFile = event.target.files?.[0] || null;

                              if (selectedFile) {
                                const isPdf =
                                  selectedFile.type === 'application/pdf' ||
                                  /\.pdf$/i.test(selectedFile.name);

                                if (!isPdf) {
                                  setMessage('Solo se permiten archivos PDF.');
                                  event.target.value = '';
                                  return;
                                }
                              }

                              setMessage('');
                              actualizarCotizacion(cotizacionIndex, { archivo: selectedFile });
                            }}
                          />
                          <span className={styles.fileUploadIcon} aria-hidden="true">↑</span>
                          <span className={styles.fileUploadText}>Adjuntar PDF</span>
                        </label>
                      )}
                    </div>

                    <div className={styles.formGroup}>
                      <label className={styles.label}>Valor</label>
                      <div className={styles.moneyInput}>
                        <span>$</span>
                        <input
                          className={styles.input}
                          type="number"
                          min="0"
                          step="0.01"
                          placeholder="0.00"
                          value={cotizacion.valorTotal}
                          aria-label={`Valor de cotización ${cotizacionIndex + 1}`}
                          onChange={(event: React.ChangeEvent<HTMLInputElement>) =>
                            actualizarCotizacion(cotizacionIndex, { valorTotal: event.target.value })
                          }
                        />
                      </div>
                    </div>
                  </div>

                  <div className={styles.quotationFooter}>
                    <span>Total</span>
                    <strong>{formatCurrency(Number(cotizacion.valorTotal) || 0)}</strong>
                  </div>
                </div>
              ))}

              {cotizaciones.length < 3 && (
                <button
                  type="button"
                  className={styles.addQuotationTile}
                  onClick={agregarCotizacion}
                  aria-label="Agregar cotización"
                >
                  <span className={styles.addQuotationIcon}>+</span>
                  <strong>Agregar cotización</strong>
                  <small>{cotizaciones.length} / 3</small>
                </button>
              )}
            </div>
          </section>

          <div className={`${styles.stickyActions} ${styles.compactStickyActions}`}>
            <div className={styles.actions}>
              <button type="button" className={styles.secondaryButton} disabled={saving} onClick={() => { void saveRequerimiento(false); }}>
                {saving
                  ? 'Guardando...'
                  : editingDraftId
                    ? 'Actualizar borrador'
                    : 'Guardar borrador'}
              </button>
              <button type="button" className={styles.primaryButton} disabled={saving} onClick={() => { void saveRequerimiento(true); }}>
                {saving ? 'Procesando...' : 'Enviar a aprobación →'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ÓRDENES DE COMPRA - RESPONSABLE OC */}

      {activeView === 'ordenes' && user?.role === 'ResponsableOC' && (
        selectedOrdenCompra ? (
          <div className={styles.detailWorkspace}>
            <div className={styles.detailTopBar}>
              <button
                type="button"
                className={styles.backButton}
                disabled={procesandoOrdenCompra}
                onClick={cerrarOrdenCompra}
              >
                ← Órdenes de compra
              </button>

              <div className={styles.detailSectionHeading}>
                <div className={styles.detailTitleRow}>
                  <h2 className={styles.pageTitle}>
                    {selectedRequerimiento
                      ? `Requerimiento #${selectedRequerimiento.Id}`
                      : (selectedOrdenCompra.Title || `OC #${selectedOrdenCompra.Id}`)}
                  </h2>
                  <span className={getEstadoClass(
                    selectedOrdenCompra.EstadoOC === 'Generada'
                      ? 'Aprobado'
                      : selectedOrdenCompra.EstadoOC === 'Rechazada'
                        ? 'Rechazado'
                        : 'Enviado Aprobacion'
                  )}>
                    {selectedOrdenCompra.EstadoOC}
                  </span>
                </div>
                <span className={styles.roleAccentBadge}>
                  {selectedOrdenCompra.TipoDocumento}
                </span>
              </div>
            </div>

            {selectedRequerimiento && (
              <>
                <section className={`${styles.detailCard} ${styles.unifiedRequestCard}`}>
                  <div className={styles.readOnlyFieldsGrid}>
                    <div className={styles.readOnlyField}>
                      <small>Área</small>
                      <strong>{selectedRequerimiento.Area || 'Sin área'}</strong>
                    </div>

                    <div className={styles.readOnlyField}>
                      <small>Categoría</small>
                      <strong>{getCategoriaNombre(selectedRequerimiento.CategoriaId)}</strong>
                    </div>

                    <div className={styles.readOnlyField}>
                      <small>Tipo de proceso</small>
                      <strong>
                        {selectedRequerimiento.Recurrente
                          ? 'Recurrente · Contrato'
                          : 'No recurrente · OCV'}
                      </strong>
                    </div>

                    <div className={styles.readOnlyField}>
                      <small>Solicitado por</small>
                      <strong>
                        {selectedRequerimiento.Solicitante?.Title || 'Sin solicitante'}
                      </strong>
                    </div>

                    <div className={styles.readOnlyField}>
                      <small>Fecha de solicitud</small>
                      <strong>{formatDateTime(selectedRequerimiento.Created)}</strong>
                    </div>

                    <div className={styles.readOnlyField}>
                      <small>Valor aprobado</small>
                      <strong>{formatCurrency(selectedRequerimiento.ValorTotal)}</strong>
                    </div>
                  </div>

                  <div className={styles.compactDescriptionReadOnly}>
                    <small>Descripción</small>
                    <p title={selectedRequerimiento.Descripcion}>
                      {selectedRequerimiento.Descripcion}
                    </p>
                  </div>

                  <div className={styles.compactDecisionInfo}>
                    <small>APROBACIÓN</small>
                    <span>
                      <strong>Aprobado por </strong>
                      {selectedRequerimiento.DecisionPor?.Title ||
                        selectedRequerimiento.Aprobador?.Title ||
                        'Sin información'}
                      {(selectedRequerimiento.FechaDecision ||
                        selectedRequerimiento.FechaAprobacion) && (
                        <>
                          {' · '}
                          {formatDateTime(
                            selectedRequerimiento.FechaDecision ||
                            selectedRequerimiento.FechaAprobacion
                          )}
                        </>
                      )}
                      {selectedRequerimiento.ComentarioAprobador && (
                        <> · {selectedRequerimiento.ComentarioAprobador}</>
                      )}
                    </span>
                  </div>

                  <div className={styles.compactTimeline} aria-label="Etapas del proceso">
                    {getEtapas(selectedRequerimiento).map((etapa, index, etapas) => {
                      const currentIndex = Math.max(
                        0,
                        etapas.indexOf(
                          selectedRequerimiento.EtapaActual || 'Solicitud'
                        )
                      );

                      const completed =
                        index < currentIndex ||
                        selectedRequerimiento.EtapaActual === 'Finalizado';

                      const current =
                        index === currentIndex &&
                        selectedRequerimiento.EtapaActual !== 'Finalizado';

                      const consultable =
                        index <= currentIndex ||
                        selectedRequerimiento.EtapaActual === 'Finalizado';

                      const selected = etapaConsulta === etapa;

                      return (
                        <button
                          key={etapa}
                          type="button"
                          className={`${styles.compactTimelineStep} ${styles.timelineStepButton} ${completed ? styles.timelineCompleted : ''} ${current ? styles.timelineCurrent : ''} ${selected ? styles.timelineSelected : ''}`}
                          aria-current={current ? 'step' : undefined}
                          aria-pressed={selected}
                          disabled={!consultable}
                          onClick={() => {
                            if (consultable) {
                              setEtapaConsulta(etapa);
                            }
                          }}
                        >
                          <div className={styles.compactTimelineMarker}>
                            {completed ? '✓' : index + 1}
                          </div>
                          <span>{getEtapaLabel(etapa)}</span>
                        </button>
                      );
                    })}
                  </div>

                  {etapaConsulta && (() => {
                    const etapaData = getStageInspection(
                      selectedRequerimiento,
                      etapaConsulta
                    );

                    return (
                      <div className={styles.stageInspection}>
                        <div className={styles.stageInspectionHeader}>
                          <div>
                            <small>DETALLE DE ETAPA</small>
                            <strong>{etapaData.title}</strong>
                          </div>
                          <span>{etapaData.status}</span>
                        </div>

                        <div className={styles.stageInspectionMeta}>
                          {etapaData.date && (
                            <span>{formatDateTime(etapaData.date)}</span>
                          )}
                          {etapaData.actor && (
                            <span>{etapaData.actor}</span>
                          )}
                        </div>

                        <div className={styles.stageInspectionDetails}>
                          {etapaData.details.map(
                            (detail: string, detailIndex: number) => (
                              <div key={`${etapaData.title}-${detailIndex}`}>
                                {detail}
                              </div>
                            )
                          )}
                        </div>
                      </div>
                    );
                  })()}
                </section>

                <section className={`${styles.formSection} ${styles.detailQuotationsSection}`}>
                  <div className={`${styles.sectionHeaderRow} ${styles.compactQuotationHeader}`}>
                    <h3>Cotización aprobada</h3>
                  </div>

                  {loadingDetalle ? (
                    <div className={styles.emptyState}>Cargando cotización...</div>
                  ) : cotizacionesDetalle.filter(
                        (cotizacion) => cotizacion.Seleccionada
                      ).length === 0 ? (
                    <div className={styles.emptyState}>
                      No se encontró la cotización aprobada.
                    </div>
                  ) : (
                    <div className={styles.approvalQuotationGrid}>
                      {cotizacionesDetalle
                        .filter((cotizacion) => cotizacion.Seleccionada)
                        .map((cotizacion: ICotizacionDetalle) => (
                          <div
                            key={cotizacion.Id}
                            className={`${styles.approvalQuotationCard} ${styles.approvalQuotationCardSelected}`}
                          >
                            <div className={styles.readOnlyQuotationHeader}>
                              <span className={styles.quotationLabel}>
                                COTIZACIÓN APROBADA
                              </span>
                              <div className={styles.readOnlyQuotationHeaderMeta}>
                                {cotizacion.CodigoCotizacionInterno && (
                                  <small>{cotizacion.CodigoCotizacionInterno}</small>
                                )}
                                <span className={styles.selectionCheck}>✓</span>
                              </div>
                            </div>

                            <div className={styles.readOnlyQuotationBody}>
                              <div className={styles.readOnlyQuotationField}>
                                <small>PDF</small>
                                {cotizacion.Attachments.length > 0 ? (
                                  cotizacion.Attachments.map((archivo) => (
                                    <div
                                      key={archivo.ServerRelativeUrl}
                                      className={styles.readOnlyPdfChip}
                                    >
                                      <span className={styles.pdfFileIcon}>PDF</span>
                                      <span
                                        className={styles.pdfFileName}
                                        title={archivo.FileName}
                                      >
                                        {archivo.FileName}
                                      </span>
                                      <button
                                        type="button"
                                        className={styles.previewIconButton}
                                        onClick={() => setArchivoPreview(archivo)}
                                      >
                                        Ver PDF
                                      </button>
                                    </div>
                                  ))
                                ) : (
                                  <span className={styles.noFileText}>
                                    Sin documento adjunto
                                  </span>
                                )}
                              </div>

                              <div className={styles.readOnlyQuotationField}>
                                <small>Valor aprobado</small>
                                <strong className={styles.readOnlyQuotationValue}>
                                  {formatCurrency(cotizacion.ValorTotal)}
                                </strong>
                              </div>
                            </div>
                          </div>
                        ))}
                    </div>
                  )}
                </section>
              </>
            )}

            <section className={`${styles.detailCard} ${styles.unifiedRequestCard}`}>
              <div className={styles.readOnlyFieldsGrid}>
                <div className={styles.readOnlyField}>
                  <small>Requerimiento</small>
                  <strong>
                    {selectedOrdenCompra.RequerimientoId
                      ? `#${selectedOrdenCompra.RequerimientoId}`
                      : 'Sin requerimiento'}
                  </strong>
                </div>

                <div className={styles.readOnlyField}>
                  <small>Tipo de gestión</small>
                  <strong>
                    {selectedOrdenCompra.TipoOrden === 'GenerarOC'
                      ? 'Generación automática'
                      : 'Ingreso desde ERP'}
                  </strong>
                </div>

                <div className={styles.readOnlyField}>
                  <small>Solicitado por</small>
                  <strong>
                    {selectedOrdenCompra.SolicitadoPor?.Title || 'Sin solicitante'}
                  </strong>
                </div>

                <div className={styles.readOnlyField}>
                  <small>Fecha solicitud</small>
                  <strong>
                    {selectedOrdenCompra.FechaSolicitud
                      ? formatDateTime(selectedOrdenCompra.FechaSolicitud)
                      : 'Sin fecha'}
                  </strong>
                </div>

                <div className={styles.readOnlyField}>
                  <small>Número OC</small>
                  <strong>
                    {selectedOrdenCompra.NumeroOC || 'Pendiente'}
                  </strong>
                </div>

                <div className={styles.readOnlyField}>
                  <small>Fecha generación</small>
                  <strong>
                    {selectedOrdenCompra.FechaGeneracion
                      ? formatDateTime(selectedOrdenCompra.FechaGeneracion)
                      : 'Pendiente'}
                  </strong>
                </div>
              </div>

              {selectedOrdenCompra.TipoOrden === 'GenerarOC' ? (
                <div className={styles.compactDecisionInfo}>
                  <small>GENERACIÓN AUTOMÁTICA</small>
                  <span>
                    Esta orden fue generada automáticamente al aprobar el requerimiento.
                    {' '}El formato actual del número es provisional hasta que negocio confirme
                    la estructura definitiva.
                  </span>
                </div>
              ) : selectedOrdenCompra.EstadoOC === 'Generada' ? (
                <div className={styles.compactDecisionInfo}>
                  <small>ORDEN REGISTRADA</small>
                  <span>
                    El número de orden emitido por el ERP ya fue registrado y el
                    requerimiento avanzó a Facturación.
                  </span>
                </div>
              ) : (
                <div className={styles.compactDecisionInfo}>
                  <small>PENDIENTE DE INGRESAR OC</small>
                  <span>
                    Gestiona la orden en el ERP y registra aquí el número
                    generado. No se genera ningún PDF o comprobante desde esta aplicación.
                  </span>
                </div>
              )}
            </section>

            {selectedOrdenCompra.TipoOrden === 'SolicitarOC' &&
              selectedOrdenCompra.EstadoOC !== 'Generada' && (
              <section className={`${styles.detailCard} ${styles.compactDecisionCard}`}>
                <div className={styles.compactDecisionHeading}>
                  <div>
                    <span className={styles.eyebrow}>REGISTRO DE ORDEN</span>
                    <h3>Ingresar orden de compra</h3>
                  </div>
                  <span>
                    Al guardar, la orden cambiará a Generada y el requerimiento
                    avanzará a Facturación.
                  </span>
                </div>

                <div className={styles.formGroup}>
                  <label className={styles.label} htmlFor="numero-oc-erp">
                    Número de orden de compra
                  </label>
                  <input
                    id="numero-oc-erp"
                    className={styles.input}
                    type="text"
                    value={numeroOCInput}
                    disabled={procesandoOrdenCompra}
                    placeholder="Ej. 4500012876"
                    onChange={(event: React.ChangeEvent<HTMLInputElement>) =>
                      setNumeroOCInput(event.target.value)
                    }
                  />
                </div>

                <div className={styles.formGroup}>
                  <label className={styles.label} htmlFor="observaciones-oc">
                    Observaciones
                  </label>
                  <textarea
                    id="observaciones-oc"
                    className={styles.textarea}
                    rows={3}
                    value={observacionesOCInput}
                    disabled={procesandoOrdenCompra}
                    placeholder="Opcional"
                    onChange={(event: React.ChangeEvent<HTMLTextAreaElement>) =>
                      setObservacionesOCInput(event.target.value)
                    }
                  />
                </div>

                <div className={styles.actions}>
                  <button
                    type="button"
                    className={styles.secondaryButton}
                    disabled={procesandoOrdenCompra}
                    onClick={cerrarOrdenCompra}
                  >
                    Cancelar
                  </button>
                  <button
                    type="button"
                    className={styles.primaryButton}
                    disabled={procesandoOrdenCompra || !numeroOCInput.trim()}
                    onClick={() => { void registrarNumeroOrdenCompra(); }}
                  >
                    {procesandoOrdenCompra
                      ? 'Registrando...'
                      : 'Registrar OC →'}
                  </button>
                </div>
              </section>
            )}

            {selectedRequerimiento &&
              selectedOrdenCompra.EstadoOC === 'Generada' &&
              (selectedRequerimiento.EtapaActual === 'Facturacion' ||
                selectedRequerimiento.EtapaActual === 'Finalizado') && (
              <section className={`${styles.detailCard} ${styles.compactDecisionCard}`}>
                <div className={styles.compactDecisionHeading}>
                  <div>
                    <span className={styles.eyebrow}>FACTURACIÓN</span>
                    <h3>Factura asociada a la orden de compra</h3>
                  </div>
                  <span>
                    {selectedRequerimiento.EtapaActual === 'Finalizado'
                      ? 'La factura ya fue registrada y el requerimiento está finalizado.'
                      : 'Adjunta el PDF de la factura para cerrar el requerimiento.'}
                  </span>
                </div>

                {selectedOrdenCompra.AttachmentFiles?.length > 0 ? (
                  <div className={styles.readOnlyQuotationField}>
                    <small>PDF de factura</small>
                    {selectedOrdenCompra.AttachmentFiles.map((archivo) => (
                      <div
                        key={archivo.ServerRelativeUrl}
                        className={styles.readOnlyPdfChip}
                      >
                        <span className={styles.pdfFileIcon} aria-hidden="true">
                          PDF
                        </span>
                        <span
                          className={styles.pdfFileName}
                          title={archivo.FileName}
                        >
                          {archivo.FileName}
                        </span>
                        <button
                          type="button"
                          className={styles.previewIconButton}
                          title="Visualizar factura"
                          aria-label={`Visualizar ${archivo.FileName}`}
                          onClick={() => setArchivoPreview(archivo)}
                        >
                          Ver PDF
                        </button>
                      </div>
                    ))}
                  </div>
                ) : selectedRequerimiento.EtapaActual === 'Facturacion' ? (
                  <div className={styles.formGroup}>
                    <label className={styles.label} htmlFor="factura-pdf">
                      PDF de la factura
                    </label>

                    {facturaPdf ? (
                      <div className={styles.pdfFileChip}>
                        <span className={styles.pdfFileIcon} aria-hidden="true">
                          PDF
                        </span>
                        <span
                          className={styles.pdfFileName}
                          title={facturaPdf.name}
                        >
                          {facturaPdf.name}
                        </span>
                        <button
                          type="button"
                          className={styles.pdfFileRemove}
                          disabled={procesandoFactura}
                          onClick={() => setFacturaPdf(null)}
                          aria-label="Quitar factura seleccionada"
                          title="Quitar archivo"
                        >
                          ×
                        </button>
                      </div>
                    ) : (
                      <label className={styles.fileUpload}>
                        <input
                          id="factura-pdf"
                          type="file"
                          accept="application/pdf,.pdf"
                          disabled={procesandoFactura}
                          aria-label="Adjuntar PDF de factura"
                          onChange={(event: React.ChangeEvent<HTMLInputElement>) => {
                            const archivo = event.target.files?.[0] || null;

                            if (archivo) {
                              const esPdf =
                                archivo.type === 'application/pdf' ||
                                /\.pdf$/i.test(archivo.name);

                              if (!esPdf) {
                                setMessage('La factura debe ser un archivo PDF.');
                                event.target.value = '';
                                return;
                              }
                            }

                            setMessage('');
                            setFacturaPdf(archivo);
                          }}
                        />
                        <span className={styles.fileUploadIcon} aria-hidden="true">
                          ↑
                        </span>
                        <span className={styles.fileUploadText}>
                          Adjuntar factura PDF
                        </span>
                      </label>
                    )}
                  </div>
                ) : null}

                {selectedRequerimiento.EtapaActual === 'Facturacion' && (
                  <div className={styles.actions}>
                    <button
                      type="button"
                      className={styles.primaryButton}
                      disabled={
                        procesandoFactura ||
                        (!selectedOrdenCompra.AttachmentFiles?.length && !facturaPdf)
                      }
                      onClick={() => {
                        registrarFacturaYFinalizar().catch((error: unknown) => {
                          setMessage(getErrorMessage(error));
                        });
                      }}
                    >
                      {procesandoFactura
                        ? 'Procesando...'
                        : selectedOrdenCompra.AttachmentFiles?.length
                          ? 'Finalizar requerimiento →'
                          : 'Guardar factura y finalizar →'}
                    </button>
                  </div>
                )}
              </section>
            )}

            {selectedOrdenCompra.Observaciones && (
              <section className={styles.detailCard}>
                <div className={styles.compactDescriptionReadOnly}>
                  <small>Observaciones</small>
                  <p>{selectedOrdenCompra.Observaciones}</p>
                </div>
              </section>
            )}

            {selectedRequerimiento && (
              <section className={`${styles.detailCard} ${styles.historyCard}`}>
                <div className={styles.historyHeader}>
                  <div>
                    <span className={styles.eyebrow}>TRAZABILIDAD</span>
                    <h3>Historial del requerimiento</h3>
                  </div>
                  <small>
                    Solicitud, aprobación y orden de compra en una sola vista.
                  </small>
                </div>

                <div className={styles.historyList}>
                  {getHistoryEvents(selectedRequerimiento)
                    .slice()
                    .reverse()
                    .map((event: IHistoryEvent) => (
                      <div
                        key={event.key}
                        className={styles.historyItem}
                      >
                        <div className={styles.historyRail}>
                          <span className={styles.historyDot} />
                        </div>

                        <div className={styles.historyContent}>
                          <div className={styles.historyItemHeader}>
                            <strong>{event.title}</strong>
                            <span>{formatDateTime(event.date)}</span>
                          </div>

                          {event.actor && (
                            <small>Por {event.actor}</small>
                          )}

                          {event.detail && (
                            <p>{event.detail}</p>
                          )}
                        </div>
                      </div>
                    ))}
                </div>
              </section>
            )}
          </div>
        ) : (
          <div className={styles.requestsWorkspace}>
            <div className={styles.listHeaderBand}>
              <div className={styles.pageIntro}>
                <div>
                  <span className={styles.eyebrow}>GESTIÓN OPERATIVA</span>
                  <h2 className={styles.pageTitle}>Órdenes de compra</h2>
                </div>
              </div>

              <div className={styles.kpiGrid}>
                <div className={styles.kpiCard}>
                  <span>Total</span>
                  <strong>{ordenesCompra.length}</strong>
                </div>
                <div className={styles.kpiCard}>
                  <span>Pendientes de ingresar</span>
                  <strong>
                    {ordenesCompra.filter(
                      (orden) =>
                        orden.TipoOrden === 'SolicitarOC' &&
                        orden.EstadoOC === 'Pendiente'
                    ).length}
                  </strong>
                </div>
                <div className={styles.kpiCard}>
                  <span>Automáticas</span>
                  <strong>
                    {ordenesCompra.filter(
                      (orden) =>
                        orden.TipoOrden === 'GenerarOC' &&
                        orden.EstadoOC === 'Generada'
                    ).length}
                  </strong>
                </div>
                <div className={styles.kpiCard}>
                  <span>OC registradas</span>
                  <strong>
                    {ordenesCompra.filter(
                      (orden) => orden.EstadoOC === 'Generada'
                    ).length}
                  </strong>
                </div>
              </div>
            </div>

            <div className={styles.tableToolbar}>
              <div className={styles.searchBox}>
                <span>⌕</span>
                <input
                  value={searchTerm}
                  onChange={(event: React.ChangeEvent<HTMLInputElement>) =>
                    setSearchTerm(event.target.value)
                  }
                  placeholder="Buscar por OC, requerimiento, tipo, estado o número..."
                />
              </div>
              <span className={styles.resultCount}>
                {filteredOrdenesCompra.length} resultados
              </span>
            </div>

            <div className={styles.requestsSection}>
              {loadingOrdenesCompra ? (
                <div className={styles.emptyState}>
                  Cargando órdenes de compra...
                </div>
              ) : filteredOrdenesCompra.length === 0 ? (
                <div className={styles.emptyState}>
                  No tienes órdenes de compra asignadas.
                </div>
              ) : (
                <>
                  <table className={styles.requestsTable}>
                    <thead>
                      <tr>
                        <th>OC</th>
                        <th>Requerimiento</th>
                        <th>Tipo</th>
                        <th>Estado</th>
                        <th>Número OC</th>
                        <th>Fecha</th>
                        <th />
                      </tr>
                    </thead>
                    <tbody>
                      {filteredOrdenesCompra.map((orden: IOrdenCompra) => (
                        <tr
                          key={orden.Id}
                          className={styles.clickableRow}
                          onClick={() => {
                            abrirOrdenCompra(orden).catch((error: unknown) => {
                              setMessage(getErrorMessage(error));
                            });
                          }}
                        >
                          <td>
                            <strong>
                              {orden.Title || `OC #${orden.Id}`}
                            </strong>
                          </td>

                          <td>
                            <strong>
                              {orden.RequerimientoId
                                ? `#${orden.RequerimientoId}`
                                : 'Sin requerimiento'}
                            </strong>
                            <br />
                            <small>
                              {orden.SolicitadoPor?.Title || 'Sin solicitante'}
                            </small>
                          </td>

                          <td>
                            <div className={styles.tableStatusCell}>
                              <span className={styles.stageBadge}>
                                {orden.TipoDocumento || 'Sin tipo'}
                              </span>
                              <small className={styles.decisionText}>
                                {orden.TipoOrden === 'GenerarOC'
                                  ? 'Automática'
                                  : 'Ingresar desde ERP'}
                              </small>
                            </div>
                          </td>

                          <td>
                            <span className={getEstadoClass(
                              orden.EstadoOC === 'Generada'
                                ? 'Aprobado'
                                : orden.EstadoOC === 'Rechazada'
                                  ? 'Rechazado'
                                  : 'Enviado Aprobacion'
                            )}>
                              {orden.EstadoOC}
                            </span>
                          </td>

                          <td>
                            <strong>
                              {orden.NumeroOC ||
                                (orden.TipoOrden === 'SolicitarOC'
                                  ? 'Pendiente ERP'
                                  : 'Pendiente')}
                            </strong>
                          </td>

                          <td>
                            {orden.FechaGeneracion
                              ? new Date(orden.FechaGeneracion).toLocaleDateString('es-EC')
                              : orden.FechaSolicitud
                                ? new Date(orden.FechaSolicitud).toLocaleDateString('es-EC')
                                : 'Sin fecha'}
                          </td>

                          <td>
                            <button
                              type="button"
                              className={styles.viewButton}
                              onClick={(event) => {
                                event.stopPropagation();
                                abrirOrdenCompra(orden).catch((error: unknown) => {
                                  setMessage(getErrorMessage(error));
                                });
                              }}
                            >
                              {orden.TipoOrden === 'SolicitarOC' &&
                                orden.EstadoOC !== 'Generada'
                                ? 'Ingresar OC →'
                                : 'Ver expediente →'}
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>

                  <div className={styles.mobileRequestsList}>
                    {filteredOrdenesCompra.map((orden: IOrdenCompra) => (
                      <button
                        key={`mobile-oc-${orden.Id}`}
                        type="button"
                        className={styles.mobileRequestCard}
                        onClick={() => {
                            abrirOrdenCompra(orden).catch((error: unknown) => {
                              setMessage(getErrorMessage(error));
                            });
                          }}
                      >
                        <div className={styles.mobileRequestCardHeader}>
                          <div className={styles.mobileRequestIdBlock}>
                            <strong>
                              {orden.Title || `OC #${orden.Id}`}
                            </strong>
                            <span>
                              {orden.FechaGeneracion
                                ? new Date(orden.FechaGeneracion).toLocaleDateString('es-EC')
                                : orden.FechaSolicitud
                                  ? new Date(orden.FechaSolicitud).toLocaleDateString('es-EC')
                                  : 'Sin fecha'}
                            </span>
                          </div>

                          <span className={getEstadoClass(
                            orden.EstadoOC === 'Generada'
                              ? 'Aprobado'
                              : orden.EstadoOC === 'Rechazada'
                                ? 'Rechazado'
                                : 'Enviado Aprobacion'
                          )}>
                            {orden.EstadoOC}
                          </span>
                        </div>

                        <div className={styles.mobileRequestDescription}>
                          {orden.RequerimientoId
                            ? `Requerimiento #${orden.RequerimientoId}`
                            : 'Sin requerimiento asociado'}
                        </div>

                        <div className={styles.mobileRequestMeta}>
                          <div>
                            <small>Tipo</small>
                            <strong>{orden.TipoDocumento || 'Sin tipo'}</strong>
                          </div>
                          <div>
                            <small>Número OC</small>
                            <strong>
                              {orden.NumeroOC ||
                                (orden.TipoOrden === 'SolicitarOC'
                                  ? 'Pendiente ERP'
                                  : 'Pendiente')}
                            </strong>
                          </div>
                        </div>

                        <div className={styles.mobileRequestFooter}>
                          <div className={styles.mobileRequestStage}>
                            <span className={styles.stageBadge}>
                              {orden.TipoOrden === 'GenerarOC'
                                ? 'Automática'
                                : 'ERP'}
                            </span>
                            <small>
                              {orden.SolicitadoPor?.Title || 'Sin solicitante'}
                            </small>
                          </div>

                          <span className={styles.mobileRequestOpen}>
                            {orden.TipoOrden === 'SolicitarOC' &&
                              orden.EstadoOC !== 'Generada'
                              ? 'Ingresar OC →'
                              : 'Ver expediente →'}
                          </span>
                        </div>
                      </button>
                    ))}
                  </div>
                </>
              )}
            </div>
          </div>
        )
      )}

      {/* LISTADO / EXPEDIENTE */}

      {activeView !== 'nueva' && activeView !== 'ordenes' && user?.role !== 'SinRol' && (
        selectedRequerimiento ? (
          <div className={`${styles.detailWorkspace} ${user?.role === 'Aprobador' ? styles.approverDetail : styles.requesterDetail}`}>
            <div className={styles.detailTopBar}>
              <button
                type="button"
                className={styles.backButton}
                onClick={() => {
                  setSelectedRequerimiento(null);
                  setOrdenCompraDetalle(null);
                  setCotizacionesDetalle([]);
                  setEtapaConsulta(null);
                }}
              >
                ← Solicitudes
              </button>

              <div className={styles.detailSectionHeading}>
                <div className={styles.detailTitleRow}>
                  <h2 className={styles.pageTitle}>Requerimiento #{selectedRequerimiento.Id}</h2>
                  <span className={getEstadoClass(selectedRequerimiento.Estado)}>
                    {selectedRequerimiento.Estado}
                  </span>
                </div>
                {(user?.role === 'Aprobador' || user?.role === 'Admin') && (
                  <span className={styles.roleAccentBadge}>Aprobador</span>
                )}
              </div>
            </div>

            <section className={`${styles.detailCard} ${styles.unifiedRequestCard}`}>
              <div className={styles.readOnlyFieldsGrid}>
                <div className={styles.readOnlyField}>
                  <small>Categoría</small>
                  <strong>{getCategoriaNombre(selectedRequerimiento.CategoriaId)}</strong>
                </div>
                <div className={styles.readOnlyField}>
                  <small>Tipo de proceso</small>
                  <strong>{selectedRequerimiento.Recurrente ? 'Recurrente' : 'No recurrente'}</strong>
                </div>
                <div className={styles.readOnlyField}>
                  <small>Solicitado por</small>
                  <strong>{selectedRequerimiento.Solicitante?.Title || 'Sin solicitante'}</strong>
                </div>
                <div className={styles.readOnlyField}>
                  <small>Fecha</small>
                  <strong>{new Date(selectedRequerimiento.Created).toLocaleDateString('es-EC')}</strong>
                </div>
                <div className={styles.readOnlyField}>
                  <small>Valor promedio</small>
                  <strong>
                    {formatCurrency(
                      selectedRequerimiento.ValorPromedio ??
                      selectedRequerimiento.ValorTotal ??
                      0
                    )}
                  </strong>
                </div>
                {selectedRequerimiento.Estado === 'Aprobado' &&
                  selectedRequerimiento.ValorPromedio !== undefined &&
                  selectedRequerimiento.ValorPromedio !== null && (
                  <div className={styles.readOnlyField}>
                    <small>Valor aprobado</small>
                    <strong>{formatCurrency(selectedRequerimiento.ValorTotal)}</strong>
                  </div>
                )}
              </div>

              <div className={styles.compactDescriptionReadOnly}>
                <small>Descripción</small>
                <p title={selectedRequerimiento.Descripcion}>{selectedRequerimiento.Descripcion}</p>
              </div>

              <div className={styles.detailMetaRow}>
                <div>
                  <small>Etapa actual</small>
                  <span className={styles.stageBadge}>{getEtapaLabel(selectedRequerimiento.EtapaActual)}</span>
                </div>
                <div>
                  <small>Estado</small>
                  <span className={getEstadoClass(selectedRequerimiento.Estado)}>{selectedRequerimiento.Estado}</span>
                </div>
              </div>

              {(selectedRequerimiento.Estado === 'Aprobado' ||
                selectedRequerimiento.Estado === 'Rechazado') && (
                <div className={styles.compactDecisionInfo}>
                  <small>Decisión del aprobador</small>
                  <span>
                    <strong>
                      {selectedRequerimiento.Estado === 'Aprobado' ? 'Aprobado por' : 'Rechazado por'}
                    </strong>{' '}
                    {selectedRequerimiento.DecisionPor?.Title || 'Sin información'}
                    {(selectedRequerimiento.FechaDecision || selectedRequerimiento.FechaAprobacion) && (
                      <> · {new Date(
                        selectedRequerimiento.FechaDecision ||
                        selectedRequerimiento.FechaAprobacion ||
                        ''
                      ).toLocaleDateString('es-EC')}</>
                    )}
                    {selectedRequerimiento.ComentarioAprobador && (
                      <> · {selectedRequerimiento.ComentarioAprobador}</>
                    )}
                  </span>
                </div>
              )}

              {ordenCompraDetalle &&
                selectedRequerimiento.Estado === 'Aprobado' && (
                <div className={styles.compactDecisionInfo}>
                  <small>ORDEN DE COMPRA</small>
                  <span>
                    <strong>
                      {ordenCompraDetalle.EstadoOC === 'Generada'
                        ? 'OC registrada'
                        : 'OC pendiente'}
                    </strong>
                    {' · '}
                    {ordenCompraDetalle.TipoDocumento}
                    {ordenCompraDetalle.NumeroOC && (
                      <> · Número: {ordenCompraDetalle.NumeroOC}</>
                    )}
                    {ordenCompraDetalle.FechaGeneracion && (
                      <>
                        {' · '}
                        {formatDateTime(ordenCompraDetalle.FechaGeneracion)}
                      </>
                    )}
                    {!ordenCompraDetalle.FechaGeneracion &&
                      ordenCompraDetalle.ResponsableOC?.Title && (
                      <>
                        {' · '}Responsable:{' '}
                        {ordenCompraDetalle.ResponsableOC.Title}
                      </>
                    )}
                  </span>
                </div>
              )}

              <div className={styles.compactTimeline} aria-label="Etapas del proceso">
                {getEtapas(selectedRequerimiento).map((etapa, index, etapas) => {
                  const currentIndex = Math.max(
                    0,
                    etapas.indexOf(
                      selectedRequerimiento.EtapaActual || 'Solicitud'
                    )
                  );
                  const completed =
                    index < currentIndex ||
                    selectedRequerimiento.EtapaActual === 'Finalizado';
                  const current =
                    index === currentIndex &&
                    selectedRequerimiento.EtapaActual !== 'Finalizado';
                  const consultable =
                    index <= currentIndex ||
                    selectedRequerimiento.EtapaActual === 'Finalizado';
                  const selected =
                    etapaConsulta === etapa;

                  return (
                    <button
                      key={etapa}
                      type="button"
                      className={`${styles.compactTimelineStep} ${styles.timelineStepButton} ${completed ? styles.timelineCompleted : ''} ${current ? styles.timelineCurrent : ''} ${selected ? styles.timelineSelected : ''}`}
                      aria-current={current ? 'step' : undefined}
                      aria-pressed={selected}
                      disabled={!consultable}
                      onClick={() => {
                        if (consultable) {
                          setEtapaConsulta(etapa);
                        }
                      }}
                    >
                      <div className={styles.compactTimelineMarker}>
                        {completed ? '✓' : index + 1}
                      </div>
                      <span>{getEtapaLabel(etapa)}</span>
                    </button>
                  );
                })}
              </div>

              {etapaConsulta && (() => {
                const etapaData = getStageInspection(
                  selectedRequerimiento,
                  etapaConsulta
                );

                return (
                  <div className={styles.stageInspection}>
                    <div className={styles.stageInspectionHeader}>
                      <div>
                        <small>DETALLE DE ETAPA</small>
                        <strong>{etapaData.title}</strong>
                      </div>
                      <span>{etapaData.status}</span>
                    </div>

                    <div className={styles.stageInspectionMeta}>
                      {etapaData.date && (
                        <span>{formatDateTime(etapaData.date)}</span>
                      )}
                      {etapaData.actor && (
                        <span>{etapaData.actor}</span>
                      )}
                    </div>

                    <div className={styles.stageInspectionDetails}>
                      {etapaData.details.map(
                        (detail: string, detailIndex: number) => (
                          <div key={`${etapaData.title}-${detailIndex}`}>
                            {detail}
                          </div>
                        )
                      )}
                    </div>
                  </div>
                );
              })()}
            </section>

            {ordenCompraDetalle && (ordenCompraDetalle.AttachmentFiles?.length ?? 0) > 0 && (
              <section className={`${styles.detailCard} ${styles.compactDecisionCard}`}>
                <div className={styles.compactDecisionHeading}>
                  <div>
                    <span className={styles.eyebrow}>FACTURACIÓN</span>
                    <h3>Factura asociada</h3>
                  </div>
                  <span>
                    Documento asociado a la orden de compra{' '}
                    {ordenCompraDetalle.NumeroOC || ''}.
                  </span>
                </div>

                <div className={styles.readOnlyQuotationField}>
                  <small>PDF de factura</small>
                  {(ordenCompraDetalle.AttachmentFiles || []).map((archivo) => (
                    <div
                      key={archivo.ServerRelativeUrl}
                      className={styles.readOnlyPdfChip}
                    >
                      <span className={styles.pdfFileIcon} aria-hidden="true">
                        PDF
                      </span>
                      <span
                        className={styles.pdfFileName}
                        title={archivo.FileName}
                      >
                        {archivo.FileName}
                      </span>
                      <button
                        type="button"
                        className={styles.previewIconButton}
                        title="Visualizar factura"
                        aria-label={`Visualizar ${archivo.FileName}`}
                        onClick={() => setArchivoPreview(archivo)}
                      >
                        Ver PDF
                      </button>
                    </div>
                  ))}
                </div>
              </section>
            )}

            <section className={`${styles.formSection} ${styles.detailQuotationsSection}`}>
              <div className={`${styles.sectionHeaderRow} ${styles.compactQuotationHeader}`}>
                <h3>Cotizaciones</h3>
                <span className={styles.counterBadge}>{cotizacionesDetalle.length} / 3</span>
              </div>

              {loadingDetalle ? (
                <div className={styles.detailQuotationSkeletonGrid} aria-busy="true" aria-label="Cargando cotizaciones">
                  {[0, 1, 2].map((item) => <div key={item} className={styles.detailQuotationSkeleton} />)}
                </div>
              ) : cotizacionesDetalle.length === 0 ? (
                <div className={styles.emptyState}>Este requerimiento no tiene cotizaciones registradas.</div>
              ) : (
                <div className={styles.approvalQuotationGrid}>
                  {cotizacionesDetalle.map((cotizacion, cotizacionIndex) => {
                    const puedeSeleccionar =
                      (user?.role === 'Aprobador' || user?.role === 'Admin') &&
                      selectedRequerimiento.Estado === 'Enviado Aprobacion';
                    const seleccionada = cotizacionSeleccionadaId === cotizacion.Id;
                    const aprobada = cotizacion.Seleccionada && selectedRequerimiento.Estado === 'Aprobado';

                    return (
                      <div
                        key={cotizacion.Id}
                        className={`${styles.approvalQuotationCard} ${
                          seleccionada || aprobada ? styles.approvalQuotationCardSelected : ''
                        } ${puedeSeleccionar ? styles.approvalQuotationCardClickable : ''}`}
                        role={puedeSeleccionar ? 'button' : undefined}
                        tabIndex={puedeSeleccionar ? 0 : undefined}
                        aria-pressed={puedeSeleccionar ? seleccionada : undefined}
                        onClick={() => {
                          if (puedeSeleccionar && !procesandoDecision) {
                            setCotizacionSeleccionadaId(cotizacion.Id);
                          }
                        }}
                        onKeyDown={(event: React.KeyboardEvent<HTMLDivElement>) => {
                          if (puedeSeleccionar && (event.key === 'Enter' || event.key === ' ')) {
                            event.preventDefault();
                            setCotizacionSeleccionadaId(cotizacion.Id);
                          }
                        }}
                      >
                        <div className={styles.readOnlyQuotationHeader}>
                          <span className={styles.quotationLabel}>COTIZACIÓN {cotizacionIndex + 1}</span>
                          <div className={styles.readOnlyQuotationHeaderMeta}>
                            {cotizacion.CodigoCotizacionInterno && (
                              <small title={cotizacion.CodigoCotizacionInterno}>{cotizacion.CodigoCotizacionInterno}</small>
                            )}
                            {(seleccionada || aprobada) && (
                              <span className={styles.selectionCheck} title="Cotización seleccionada">✓</span>
                            )}
                          </div>
                        </div>

                        <div className={styles.readOnlyQuotationBody}>
                          <div className={styles.readOnlyQuotationField}>
                            <small>PDF</small>
                            {cotizacion.Attachments.length > 0 ? cotizacion.Attachments.map((archivo) => (
                              <div key={archivo.ServerRelativeUrl} className={styles.readOnlyPdfChip}>
                                <span className={styles.pdfFileIcon} aria-hidden="true">PDF</span>
                                <span className={styles.pdfFileName} title={archivo.FileName}>{archivo.FileName}</span>
                                <button
                                  type="button"
                                  className={styles.previewIconButton}
                                  title="Visualizar PDF"
                                  aria-label={`Visualizar ${archivo.FileName}`}
                                  onClick={(event) => {
                                    event.stopPropagation();
                                    setArchivoPreview(archivo);
                                  }}
                                >
                                  <svg
                                    className={styles.previewEyeIcon}
                                    viewBox="0 0 24 24"
                                    aria-hidden="true"
                                  >
                                    <path
                                      d="M2.5 12s3.5-6 9.5-6 9.5 6 9.5 6-3.5 6-9.5 6-9.5-6-9.5-6Z"
                                      fill="none"
                                      stroke="currentColor"
                                      strokeWidth="1.8"
                                      strokeLinecap="round"
                                      strokeLinejoin="round"
                                    />
                                    <circle
                                      cx="12"
                                      cy="12"
                                      r="2.7"
                                      fill="none"
                                      stroke="currentColor"
                                      strokeWidth="1.8"
                                    />
                                  </svg>
                                  <span>Ver PDF</span>
                                </button>
                              </div>
                            )) : (
                              <span className={styles.noFileText}>Sin documento adjunto</span>
                            )}
                          </div>

                          <div className={styles.readOnlyQuotationField}>
                            <small>Valor</small>
                            <strong className={styles.readOnlyQuotationValue}>{formatCurrency(cotizacion.ValorTotal)}</strong>
                          </div>
                        </div>

                        <div className={styles.readOnlyQuotationFooter}>
                          <span>Total</span>
                          <strong>{formatCurrency(cotizacion.ValorTotal)}</strong>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </section>

            {(user?.role === 'Aprobador' || user?.role === 'Admin') &&
              selectedRequerimiento.Estado === 'Enviado Aprobacion' && (
                <section className={`${styles.detailCard} ${styles.compactDecisionCard}`}>
                  <div className={styles.compactDecisionHeading}>
                    <div>
                      <span className={styles.eyebrow}>DECISIÓN DEL APROBADOR</span>
                      <h3>Aprobar o rechazar requerimiento</h3>
                    </div>
                    <span>Selecciona una cotización para aprobar. El comentario es obligatorio al rechazar.</span>
                  </div>
                  <div className={styles.formGroup}>
                    <label className={styles.label}>Comentario (obligatorio al rechazar)</label>
                    <textarea
                      className={styles.textarea}
                      rows={2}
                      value={comentarioAprobador}
                      placeholder="Escribe una observación sobre la decisión..."
                      onChange={(event: React.ChangeEvent<HTMLTextAreaElement>) =>
                        setComentarioAprobador(event.target.value)
                      }
                    />
                  </div>
                </section>
              )}

            <section className={`${styles.detailCard} ${styles.historyCard}`}>
              <div className={styles.historyHeader}>
                <div>
                  <span className={styles.eyebrow}>TRAZABILIDAD</span>
                  <h3>Historial del requerimiento</h3>
                </div>
                <small>Los eventos se acumulan conforme avanza el proceso.</small>
              </div>

              <div className={styles.historyList}>
                {getHistoryEvents(selectedRequerimiento)
                  .slice()
                  .reverse()
                  .map((event: IHistoryEvent) => (
                    <div
                      key={event.key}
                      className={styles.historyItem}
                    >
                      <div className={styles.historyRail}>
                        <span className={styles.historyDot} />
                      </div>

                      <div className={styles.historyContent}>
                        <div className={styles.historyItemHeader}>
                          <strong>{event.title}</strong>
                          <span>
                            {formatDateTime(event.date)}
                          </span>
                        </div>

                        {event.actor && (
                          <small>Por {event.actor}</small>
                        )}

                        {event.detail && (
                          <p>{event.detail}</p>
                        )}
                      </div>
                    </div>
                  ))}
              </div>
            </section>

            <div className={`${styles.stickyActions} ${styles.compactStickyActions} ${styles.detailActionBar}`}>
              <div className={styles.nextStepSummary}>
                <small>Siguiente paso</small>
                <strong>
                  {selectedRequerimiento.Estado === 'Borrador'
                    ? 'Completar solicitud'
                    : `Seguimiento: ${getEtapaLabel(selectedRequerimiento.EtapaActual)}`}
                </strong>
              </div>

              {selectedRequerimiento.Estado === 'Borrador' &&
                (
                  user?.role === 'Admin' ||
                  (
                    user?.role === 'Solicitador' &&
                    selectedRequerimiento.Solicitante?.Id === user.id
                  )
                ) && (
                  <div className={styles.actions}>
                    <button
                      type="button"
                      className={styles.primaryButton}
                      disabled={loadingDetalle}
                      onClick={() => {
                        void continuarBorrador(
                          selectedRequerimiento
                        );
                      }}
                    >
                      Continuar borrador →
                    </button>
                  </div>
                )}

              {(user?.role === 'Aprobador' || user?.role === 'Admin') &&
                selectedRequerimiento.Estado === 'Enviado Aprobacion' && (
                  <div className={styles.actions}>
                    <button
                      type="button"
                      className={styles.secondaryButton}
                      disabled={procesandoDecision}
                      onClick={() => { void decidirRequerimiento(false); }}
                    >
                      {procesandoDecision ? 'Procesando...' : 'Rechazar'}
                    </button>
                    <button
                      type="button"
                      className={styles.primaryButton}
                      disabled={procesandoDecision}
                      onClick={() => { void decidirRequerimiento(true); }}
                    >
                      {procesandoDecision ? 'Procesando...' : 'Aprobar cotización'}
                    </button>
                  </div>
                )}
            </div>
          </div>
        ) : (
          <div className={styles.requestsWorkspace}>
            <div className={styles.listHeaderBand}>
              <div className={styles.pageIntro}>
                <div>
                  <span className={styles.eyebrow}>CONTROL DE EXPEDIENTES</span>
                  <h2 className={styles.pageTitle}>{activeView === 'mis' ? 'Mis solicitudes' : 'Solicitudes'}</h2>
                </div>
              </div>

              <div className={styles.kpiGrid}>
                <div className={styles.kpiCard}><span>Total</span><strong>{requerimientos.length}</strong></div>
                <div className={styles.kpiCard}><span>En aprobación</span><strong>{requerimientos.filter(r => r.Estado === 'Enviado Aprobacion').length}</strong></div>
                <div className={styles.kpiCard}><span>Aprobados</span><strong>{requerimientos.filter(r => r.Estado === 'Aprobado').length}</strong></div>
                <div className={styles.kpiCard}><span>Borradores</span><strong>{requerimientos.filter(r => r.Estado === 'Borrador').length}</strong></div>
              </div>
            </div>

            <div className={styles.tableToolbar}>
              <div className={styles.searchBox}>
                <span>⌕</span>
                <input value={searchTerm} onChange={(e: React.ChangeEvent<HTMLInputElement>) => setSearchTerm(e.target.value)} placeholder="Buscar por ID, descripción, categoría o estado..." />
              </div>
              <span className={styles.resultCount}>{filteredRequerimientos.length} resultados</span>
            </div>

            <div className={styles.requestsSection}>
              {loadingRequerimientos ? (
                <div className={styles.emptyState}>Cargando solicitudes...</div>
              ) : filteredRequerimientos.length === 0 ? (
                <div className={styles.emptyState}>No existen solicitudes para mostrar.</div>
              ) : (
                <>
                  <table className={styles.requestsTable}>
                    <thead>
                      <tr>
                        <th>ID</th>
                        <th>Requerimiento</th>
                        <th>Categoría</th>
                        <th>Valor</th>
                        <th>Seguimiento</th>
                        <th>Fecha</th>
                        <th />
                      </tr>
                    </thead>
                    <tbody>
                      {filteredRequerimientos.map((requerimiento: IRequerimientoItem) => (
                        <tr key={requerimiento.Id} className={styles.clickableRow} onClick={() => abrirRequerimiento(requerimiento)}>
                          <td><strong>#{requerimiento.Id}</strong></td>
                          <td>
                            <div className={styles.requestTitle}>{requerimiento.Descripcion}</div>
                            <small>{requerimiento.Solicitante?.Title || 'Sin solicitante'}</small>
                          </td>
                          <td>{getCategoriaNombre(requerimiento.CategoriaId)}</td>
                          <td>
                            <strong>
                              {formatCurrency(
                                requerimiento.Estado === 'Aprobado'
                                  ? requerimiento.ValorTotal
                                  : (requerimiento.ValorPromedio ?? requerimiento.ValorTotal ?? 0)
                              )}
                            </strong>
                          </td>
                          <td>
                            <div className={styles.tableStatusCell}>
                              <span className={styles.stageBadge}>{getEtapaLabel(requerimiento.EtapaActual)}</span>
                              <span className={getEstadoClass(requerimiento.Estado)}>{requerimiento.Estado}</span>
                              <small className={styles.decisionText}>
                                {requerimiento.Estado === 'Aprobado' && requerimiento.DecisionPor
                                  ? `Por ${requerimiento.DecisionPor.Title}`
                                  : requerimiento.Estado === 'Rechazado' && requerimiento.DecisionPor
                                    ? `Por ${requerimiento.DecisionPor.Title}`
                                    : requerimiento.Estado === 'Enviado Aprobacion'
                                      ? 'Pendiente de decisión'
                                      : 'Sin enviar'}
                              </small>
                            </div>
                          </td>
                          <td>
                            {new Date(
                              (requerimiento.Estado === 'Aprobado' || requerimiento.Estado === 'Rechazado') &&
                              (requerimiento.FechaDecision || requerimiento.FechaAprobacion)
                                ? (requerimiento.FechaDecision || requerimiento.FechaAprobacion || requerimiento.Created)
                                : requerimiento.Created
                            ).toLocaleDateString('es-EC')}
                          </td>
                          <td>
                            <button type="button" className={styles.viewButton} onClick={(e) => { e.stopPropagation(); abrirRequerimiento(requerimiento); }}>
                              Ver →
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>

                  <div className={styles.mobileRequestsList}>
                    {filteredRequerimientos.map((requerimiento: IRequerimientoItem) => {
                      const fechaVisible =
                        (requerimiento.Estado === 'Aprobado' || requerimiento.Estado === 'Rechazado') &&
                        (requerimiento.FechaDecision || requerimiento.FechaAprobacion)
                          ? (requerimiento.FechaDecision || requerimiento.FechaAprobacion || requerimiento.Created)
                          : requerimiento.Created;

                      const valorVisible =
                        requerimiento.Estado === 'Aprobado'
                          ? requerimiento.ValorTotal
                          : (requerimiento.ValorPromedio ?? requerimiento.ValorTotal ?? 0);

                      return (
                        <button
                          key={`mobile-${requerimiento.Id}`}
                          type="button"
                          className={styles.mobileRequestCard}
                          onClick={() => abrirRequerimiento(requerimiento)}
                          aria-label={`Abrir requerimiento ${requerimiento.Id}`}
                        >
                          <div className={styles.mobileRequestCardHeader}>
                            <div className={styles.mobileRequestIdBlock}>
                              <strong>#{requerimiento.Id}</strong>
                              <span>{new Date(fechaVisible).toLocaleDateString('es-EC')}</span>
                            </div>
                            <span className={getEstadoClass(requerimiento.Estado)}>
                              {requerimiento.Estado}
                            </span>
                          </div>

                          <div className={styles.mobileRequestDescription}>
                            {requerimiento.Descripcion}
                          </div>

                          <div className={styles.mobileRequestMeta}>
                            <div>
                              <small>Categoría</small>
                              <strong>{getCategoriaNombre(requerimiento.CategoriaId)}</strong>
                            </div>
                            <div>
                              <small>Valor</small>
                              <strong>{formatCurrency(valorVisible)}</strong>
                            </div>
                          </div>

                          <div className={styles.mobileRequestFooter}>
                            <div className={styles.mobileRequestStage}>
                              <span className={styles.stageBadge}>
                                {getEtapaLabel(requerimiento.EtapaActual)}
                              </span>
                              <small>{requerimiento.Solicitante?.Title || 'Sin solicitante'}</small>
                            </div>
                            <span className={styles.mobileRequestOpen}>Ver detalle →</span>
                          </div>
                        </button>
                      );
                    })}
                  </div>
                </>
              )}
            </div>
          </div>
        )
      )}

      {archivoPreview && (
        <div
          className={styles.pdfModalOverlay}
          role="dialog"
          aria-modal="true"
          aria-label={`Vista previa de ${archivoPreview.FileName}`}
          onClick={() => setArchivoPreview(null)}
        >
          <div className={styles.pdfModalContent} onClick={(event) => event.stopPropagation()}>
            <div className={styles.pdfModalHeader}>
              <div>
                <span className={styles.eyebrow}>VISTA PREVIA</span>
                <strong>{archivoPreview.FileName}</strong>
              </div>
              <button
                type="button"
                className={styles.pdfModalClose}
                aria-label="Cerrar vista previa"
                title="Cerrar"
                onClick={() => setArchivoPreview(null)}
              >
                ×
              </button>
            </div>
            <iframe
              title={archivoPreview.FileName}
              src={getVistaPreviaArchivo(archivoPreview)}
              className={styles.pdfPreviewFrame}
            />
          </div>
        </div>
      )}

      {/* MENSAJES */}

      {message && (

        <div
          className={
            styles.success
          }
        >
          {message}
        </div>
      )}

    </div>
  );
};

export default Requerimientos;
