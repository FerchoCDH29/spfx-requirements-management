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
  | 'SinRol';

type AppView =
  | 'nueva'
  | 'mis'
  | 'pendientes'
  | 'aprobadas'
  | 'rechazadas'
  | 'todas';

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
  CategoriaId?: number;
  AprobadorId?: number;
  ValorPromedio?: number | null;
  ValorTotal?: number | null;
  Estado: string;
  Recurrente: boolean;
  Created: string;
  EtapaActual?: string;
  FechaAprobacion?: string;
  ComentarioAprobador?: string;

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
          `Id,Title,Descripcion,CategoriaId,AprobadorId,ValorPromedio,ValorTotal,Estado,` +
          `Recurrente,Created,EtapaActual,FechaAprobacion,ComentarioAprobador,` +
          `DecisionPor/Id,DecisionPor/Title,DecisionPor/EMail,` +
          `Solicitante/Id,Solicitante/Title,Solicitante/EMail` +
          `&$expand=Solicitante,DecisionPor` +
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

      const detalle = await getCotizacionesDetalle(requerimientoId);

      setCotizacionesDetalle(detalle);

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
      `Id,Title,Descripcion,CategoriaId,AprobadorId,ValorPromedio,ValorTotal,Estado,` +
      `Recurrente,Created,EtapaActual,FechaAprobacion,ComentarioAprobador,` +
      `DecisionPor/Id,DecisionPor/Title,DecisionPor/EMail,` +
      `Solicitante/Id,Solicitante/Title,Solicitante/EMail` +
      `&$expand=Solicitante,DecisionPor`;

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
            userRole === 'Admin' || userRole === 'Aprobador'
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
      setMessage('');

      if (
        view !== 'nueva' &&
        user
      ) {
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


  const crearOrdenCompraSiNoExiste = async (
    requerimiento: IRequerimientoItem,
    cotizacionId: number,
    solicitadoPorId: number
  ): Promise<void> => {
    const filter = encodeURIComponent(
      `RequerimientoId eq ${requerimiento.Id} and Activo eq 1`
    );

    const existingUrl =
      `${context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists/getbytitle('OrdenesCompra')/items` +
      `?$select=Id&$filter=${filter}&$top=1`;

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

    if (existingData.value?.length > 0) {
      return;
    }

    await createListItem('OrdenesCompra', {
      Title: `OC-REQ-${requerimiento.Id}`,
      RequerimientoId: requerimiento.Id,
      CotizacionId: cotizacionId,
      TipoOrden: requerimiento.Recurrente ? 'GenerarOC' : 'SolicitarOC',
      TipoDocumento: requerimiento.Recurrente ? 'Contrato' : 'OCV',
      EstadoOC: 'Pendiente',
      FechaSolicitud: new Date().toISOString(),
      SolicitadoPorId: solicitadoPorId,
      Activo: true
    });
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
          solicitadoPorId
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
                ? 'GenerarOC'
                : 'SolicitarOC'
            )
          : 'Aprobacion',
        FechaAprobacion: new Date().toISOString(),
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
            ? `Requerimiento #${requerimientoId} aprobado. Se generó la orden de compra asociada a Contrato.`
            : `Requerimiento #${requerimientoId} aprobado. Se generó la solicitud de orden de compra OCV.`
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
      GenerarOC: 'Generar orden de compra',
      SolicitarOC: 'Solicitar orden de compra',
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


  const abrirRequerimiento =
  (requerimiento: IRequerimientoItem): void => {

    setSelectedRequerimiento(requerimiento);
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

      {/* LISTADO / EXPEDIENTE */}

      {activeView !== 'nueva' && user?.role !== 'SinRol' && (
        selectedRequerimiento ? (
          <div className={`${styles.detailWorkspace} ${user?.role === 'Aprobador' ? styles.approverDetail : styles.requesterDetail}`}>
            <div className={styles.detailTopBar}>
              <button
                type="button"
                className={styles.backButton}
                onClick={() => setSelectedRequerimiento(null)}
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
                    {selectedRequerimiento.FechaAprobacion && (
                      <> · {new Date(selectedRequerimiento.FechaAprobacion).toLocaleDateString('es-EC')}</>
                    )}
                    {selectedRequerimiento.ComentarioAprobador && (
                      <> · {selectedRequerimiento.ComentarioAprobador}</>
                    )}
                  </span>
                </div>
              )}

              <div className={styles.compactTimeline} aria-label="Etapas del proceso">
                {getEtapas(selectedRequerimiento).map((etapa, index, etapas) => {
                  const currentIndex = Math.max(0, etapas.indexOf(selectedRequerimiento.EtapaActual || 'Solicitud'));
                  const completed = index < currentIndex || selectedRequerimiento.EtapaActual === 'Finalizado';
                  const current = index === currentIndex && selectedRequerimiento.EtapaActual !== 'Finalizado';
                  return (
                    <div
                      key={etapa}
                      className={`${styles.compactTimelineStep} ${completed ? styles.timelineCompleted : ''} ${current ? styles.timelineCurrent : ''}`}
                      aria-current={current ? 'step' : undefined}
                    >
                      <div className={styles.compactTimelineMarker}>{completed ? '✓' : index + 1}</div>
                      <span>{getEtapaLabel(etapa)}</span>
                    </div>
                  );
                })}
              </div>
            </section>

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
                            (requerimiento.Estado === 'Aprobado' || requerimiento.Estado === 'Rechazado') && requerimiento.FechaAprobacion
                              ? requerimiento.FechaAprobacion
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
