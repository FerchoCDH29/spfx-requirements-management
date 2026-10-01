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

interface ICotizacionForm {
  key: string;
  valorTotal: string;
  archivo: File | null;
}

interface IRequerimientoItem {
  Id: number;
  Title: string;
  Descripcion: string;
  CategoriaId?: number;
  ValorTotal: number;
  Estado: string;
  Recurrente: boolean;
  Created: string;
  EtapaActual?: string;

  Solicitante?: {
    Id: number;
    Title: string;
    EMail?: string;
  };
}

interface IAdjuntoCotizacion {
  FileName: string;
  ServerRelativeUrl: string;
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
    archivo: null
  });

  const [cotizaciones, setCotizaciones] =
    useState<ICotizacionForm[]>([createEmptyCotizacion()]);

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
          `Id,Title,Descripcion,CategoriaId,ValorTotal,Estado,` +
          `Recurrente,Created,EtapaActual,` +
          `Solicitante/Id,Solicitante/Title,Solicitante/EMail` +
          `&$expand=Solicitante` +
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

  const loadDetalleRequerimiento = async (requerimientoId: number): Promise<void> => {
    try {
      setLoadingDetalle(true);
      setCotizacionesDetalle([]);
      const cotizacionesUrl =
        `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Cotizaciones')/items` +
        `?$select=Id,Title,NumeroCotizacionProveedor,CodigoCotizacionInterno,ValorTotal,Seleccionada,` +
        `AttachmentFiles/FileName,AttachmentFiles/ServerRelativeUrl` +
        `&$expand=AttachmentFiles` +
        `&$filter=${encodeURIComponent(`RequerimientoId eq ${requerimientoId} and Activo eq 1`)}` +
        `&$orderby=Id asc`;
      const response = await context.spHttpClient.get(cotizacionesUrl, SPHttpClient.configurations.v1, {
        headers: { Accept: 'application/json;odata=nometadata' }
      });
      if (!response.ok) {
        throw new Error(`Error cargando cotizaciones. HTTP ${response.status}: ${await response.text()}`);
      }
      const data = await response.json();
      const detalle: ICotizacionDetalle[] = data.value.map((cotizacion: {
        Id: number; Title: string; NumeroCotizacionProveedor?: string;
        CodigoCotizacionInterno?: string; ValorTotal?: number; Seleccionada?: boolean;
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
            userRole === 'Admin'
          );


          /*
           * Vista inicial según rol
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

  const formatCurrency = (value: number): string =>
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

  const decidirRequerimiento = async (aprobar: boolean): Promise<void> => {
    if (!selectedRequerimiento || !user) return;

    if (user.role !== 'Aprobador' && user.role !== 'Admin') {
      setMessage('No tienes permisos para tomar esta decisión.');
      return;
    }

    if (selectedRequerimiento.Estado !== 'Enviado Aprobacion') {
      setMessage('Este requerimiento ya no se encuentra pendiente de aprobación.');
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

      if (aprobar) {
        // Garantiza que exista una sola cotización seleccionada.
        for (const cotizacion of cotizacionesDetalle) {
          await updateListItem('Cotizaciones', cotizacion.Id, {
            Seleccionada: cotizacion.Id === cotizacionSeleccionadaId
          });
        }
      }

      await updateListItem('Requerimientos', requerimientoId, {
        Estado: aprobar ? 'Aprobado' : 'Rechazado',
        FechaAprobacion: new Date().toISOString(),
        ComentarioAprobador: comentarioAprobador.trim()
      });

      setSelectedRequerimiento(null);
      setComentarioAprobador('');
      setCotizacionSeleccionadaId(null);

      await loadRequerimientos(activeView, user.id);

      setMessage(
        aprobar
          ? `Requerimiento #${requerimientoId} aprobado con la cotización seleccionada.`
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
      ? ['Solicitud', 'Aprobacion', 'Contrato', 'OrdenCompra', 'Facturacion', 'Finalizado']
      : ['Solicitud', 'Aprobacion', 'OrdenCompra', 'Facturacion', 'Finalizado'];

  const getEtapaLabel = (etapa?: string): string => {
    const labels: Record<string, string> = {
      Solicitud: 'Solicitud',
      Aprobacion: 'Aprobación',
      Contrato: 'Contrato',
      OrdenCompra: 'Orden de compra',
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
    setMessage('');

    void loadDetalleRequerimiento(
      requerimiento.Id
    );
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
          (categoria: ICategory) => categoria.Id === Number(categoriaId)
        );

        if (!categoriaSeleccionada) {
          throw new Error(
            'No se pudo obtener la categoría seleccionada.'
          );
        }

        const aprobadorId = MODO_PRUEBAS
          ? categoriaSeleccionada.AprobadorPruebaId
          : categoriaSeleccionada.AprobadorCategoriaId;

        if (enviar && !aprobadorId) {
          throw new Error(
            MODO_PRUEBAS
              ? 'La categoría seleccionada no tiene un aprobador de prueba configurado.'
              : 'La categoría seleccionada no tiene un aprobador configurado.'
          );
        }

      if (cotizaciones.length < 1 || cotizaciones.length > 3) {
        throw new Error('El requerimiento debe tener entre 1 y 3 cotizaciones.');
      }

      if (enviar) {
        cotizaciones.forEach((cotizacion, cotizacionIndex) => {
          const totalCotizacion = Number(cotizacion.valorTotal);
          if (!cotizacion.valorTotal || isNaN(totalCotizacion) || totalCotizacion <= 0) {
            throw new Error(`Ingrese un valor total válido en la cotización ${cotizacionIndex + 1}.`);
          }

          if (!cotizacion.archivo) {
            throw new Error(`Adjunte el PDF de la cotización ${cotizacionIndex + 1}.`);
          }
        });
      }

      // Valor referencial del requerimiento:
      // - 1 cotización: toma ese mismo valor.
      // - 2 o 3 cotizaciones: calcula el promedio.
      const valoresCotizaciones = cotizaciones
        .map((cotizacion) => Number(cotizacion.valorTotal))
        .filter((valor) => !isNaN(valor) && valor > 0);

      const numericValue =
        valoresCotizaciones.length > 0
          ? valoresCotizaciones.reduce((total, valor) => total + valor, 0) /
            valoresCotizaciones.length
          : 0;

      const itemId = await createListItem('Requerimientos', {
        Title: descripcion.trim().substring(0, 255),
        Descripcion: descripcion.trim(),
        SolicitanteId: user.id,
        CategoriaId: Number(categoriaId),
        AprobadorId: aprobadorId,
        Recurrente: recurrente,
        ValorTotal: numericValue,
        EtapaActual: enviar ? 'Aprobacion' : 'Solicitud',
        Estado: enviar ? 'Enviado Aprobacion' : 'Borrador'
      });

      for (let c = 0; c < cotizaciones.length; c += 1) {
        const cotizacion = cotizaciones[c];

        // En borrador solo persistimos cotizaciones que tengan algún dato ingresado.
        const tieneDatos =
          !!cotizacion.archivo ||
          !!cotizacion.valorTotal;

        if (!enviar && !tieneDatos) {
          continue;
        }

        const totalCotizacion = Number(cotizacion.valorTotal) || 0;
        const cotizacionBody: Record<string, unknown> = {
          Title: `REQ-${itemId}-COT-${c + 1}`,
          RequerimientoId: itemId,
          CodigoCotizacionInterno: `COT-REQ-${('000000' + itemId).slice(-6)}-${('00' + (c + 1)).slice(-2)}`,
          ValorTotal: totalCotizacion,
          Seleccionada: false,
          Activo: true
        };


        const cotizacionId = await createListItem('Cotizaciones', cotizacionBody);

        if (cotizacion.archivo) {
          const isPdf =
            cotizacion.archivo.type === 'application/pdf' ||
            /\.pdf$/i.test(cotizacion.archivo.name);
          if (!isPdf) {
            throw new Error(`El archivo de la cotización ${c + 1} debe ser PDF.`);
          }
          await adjuntarArchivo('Cotizaciones', cotizacionId, cotizacion.archivo);
        }

      }

      setDescripcion('');
      setCategoriaId('');
      setRecurrente(false);
      setCotizaciones([createEmptyCotizacion()]);

      if (user.role === 'Solicitador') {
        setActiveView('mis');
        await loadRequerimientos('mis', user.id);
      } else {
        setActiveView('todas');
        await loadRequerimientos('todas', user.id);
      }

      setMessage(
        enviar
          ? `Requerimiento #${itemId} enviado a aprobación correctamente.`
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
      <div className={styles.requerimientos}>
        Cargando información...
      </div>
    );
  }

  return (

    <div className={styles.requerimientos}>

      {/* BARRA SUPERIOR */}

      <div className={styles.topBar}>

        <div className={styles.brand}>
          <img
            src={inovaLogo}
            alt="Inova Solutions"
            className={styles.logo}
          />
        </div>

        <div className={styles.headerUser}>

          {user && (
            <>
              <div
                className={
                  styles.headerUserText
                }
              >
                <strong>
                  {user.displayName}
                </strong>

                <span>
                  {user.area ? `${user.area} · ${user.role}` : user.role}
                </span>
              </div>

              <img
                className={
                  styles.avatarSmall
                }
                src={
                  `${context.pageContext.web.absoluteUrl}` +
                  `/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(user.email)}`
                }
                alt={
                  user.displayName
                }
              />
            </>
          )}

        </div>

      </div>

      {/* ENCABEZADO */}

      <div className={styles.header}>

        <span
          className={
            styles.sectionLabel
          }
        >
          SOLICITUDES
        </span>

        <h1
          className={
            styles.title
          }
        >
          Gestión de Requerimientos
        </h1>

        <p
          className={
            styles.subtitle
          }
        >
          Crea, consulta y da seguimiento
          a los requerimientos desde una sola vista.
        </p>

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
          <div className={styles.pageIntro}>
            <div>
              <span className={styles.eyebrow}>NUEVO EXPEDIENTE</span>
              <h2 className={styles.pageTitle}>Crear requerimiento</h2>
              <p className={styles.pageDescription}>
                Selecciona la categoría, describe la necesidad y adjunta las cotizaciones para enviarlas a aprobación.
              </p>
            </div>
            <div className={styles.formProgress}>Solicitud · Paso inicial</div>
          </div>

          <section className={styles.formSection}>
            <div className={styles.sectionHeader}>
              <div className={styles.sectionNumber}>01</div>
              <div>
                <h3>Datos del requerimiento</h3>
                <p>Selecciona la categoría, describe la necesidad y define el tipo de proceso.</p>
              </div>
            </div>

            <div className={styles.formGroup}>
              <label className={styles.label}>Categoría</label>
              <select
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
              <label className={styles.label}>Descripción del requerimiento</label>
              <textarea
                className={styles.textarea}
                value={descripcion}
                placeholder="Describe brevemente el producto, servicio o necesidad..."
                onChange={(event: React.ChangeEvent<HTMLTextAreaElement>) =>
                  setDescripcion(event.target.value)
                }
                rows={4}
              />
            </div>

            <div className={styles.formGroup}>
              <label className={styles.label}>Tipo de proceso</label>

              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
                <button
                  type="button"
                  className={`${styles.recurrentToggle} ${!recurrente ? styles.recurrentToggleActive : ''}`}
                  onClick={() => setRecurrente(false)}
                >
                  <span className={styles.toggleDot} />
                  <span>
                    <strong>No recurrente</strong>
                    <small>Pasa directo a orden de compra</small>
                  </span>
                </button>

                <button
                  type="button"
                  className={`${styles.recurrentToggle} ${recurrente ? styles.recurrentToggleActive : ''}`}
                  onClick={() => setRecurrente(true)}
                >
                  <span className={styles.toggleDot} />
                  <span>
                    <strong>Recurrente</strong>
                    <small>Requiere contrato antes de la orden de compra</small>
                  </span>
                </button>
              </div>
            </div>
          </section>

          <section className={styles.formSection}>
            <div className={styles.sectionHeaderRow}>
              <div className={styles.sectionHeader}>
                <div className={styles.sectionNumber}>02</div>
                <div>
                  <h3>Cotizaciones</h3>
                  <p>Registra hasta tres cotizaciones con su valor total y documento PDF.</p>
                </div>
              </div>
              <span className={styles.counterBadge}>{cotizaciones.length} de 3</span>
            </div>

            <div className={styles.quotationList}>
              {cotizaciones.map((cotizacion, cotizacionIndex) => (
                <div key={cotizacion.key} className={styles.quotationCard}>
                  <div className={styles.quotationHeader}>
                    <div>
                      <span className={styles.quotationLabel}>COTIZACIÓN {String(cotizacionIndex + 1)}</span>
                      <h3>Propuesta comercial</h3>
                    </div>
                    <div className={styles.quotationAmount}>
                      <small>Total cotización</small>
                      <strong>{formatCurrency(Number(cotizacion.valorTotal) || 0)}</strong>
                    </div>
                  </div>

                  <div className={styles.formGroup}>
                    <label className={styles.label}>Documento PDF</label>
                    <label className={styles.fileUpload}>
                      <input
                        type="file"
                        accept="application/pdf,.pdf"
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
                      <span>
                        {cotizacion.archivo
                          ? `✓ ${cotizacion.archivo.name}`
                          : 'Adjuntar PDF'}
                      </span>
                    </label>
                  </div>

                  <div className={styles.formGroup}>
                    <label className={styles.label}>Valor total de la cotización</label>
                    <div className={styles.moneyInput}>
                      <span>$</span>
                      <input
                        className={styles.input}
                        type="number"
                        min="0"
                        step="0.01"
                        placeholder="0.00"
                        value={cotizacion.valorTotal}
                        onChange={(event: React.ChangeEvent<HTMLInputElement>) =>
                          actualizarCotizacion(cotizacionIndex, { valorTotal: event.target.value })
                        }
                      />
                    </div>
                  </div>

                  <div className={styles.quotationFooter}>
                    {cotizaciones.length > 1 && (
                      <button type="button" className={styles.dangerLink} onClick={() => eliminarCotizacion(cotizacionIndex)}>
                        Eliminar cotización
                      </button>
                    )}
                    <div className={styles.quotationGrandTotal}>
                      <span>Total</span>
                      <strong>{formatCurrency(Number(cotizacion.valorTotal) || 0)}</strong>
                    </div>
                  </div>
                </div>
              ))}
            </div>

            {cotizaciones.length < 3 && (
              <button type="button" className={styles.addQuotationButton} onClick={agregarCotizacion}>
                <span>+</span>
                <div><strong>Agregar otra cotización</strong><small>Puedes registrar hasta 3 propuestas</small></div>
              </button>
            )}
          </section>

          <div className={styles.stickyActions}>
            <div>
              <strong>¿Listo para continuar?</strong>
              <span>Puedes guardar el avance o enviarlo al flujo de aprobación.</span>
            </div>
            <div className={styles.actions}>
              <button type="button" className={styles.secondaryButton} disabled={saving} onClick={() => { void saveRequerimiento(false); }}>
                {saving ? 'Guardando...' : 'Guardar borrador'}
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
          <div className={styles.detailWorkspace}>
            <button type="button" className={styles.backButton} onClick={() => setSelectedRequerimiento(null)}>
              ← Volver a solicitudes
            </button>

            <div className={styles.detailHero}>
              <div>
                <div className={styles.detailTopline}>
                  <span>REQUERIMIENTO #{selectedRequerimiento.Id}</span>
                  <span className={getEstadoClass(selectedRequerimiento.Estado)}>{selectedRequerimiento.Estado}</span>
                </div>
                <h2>{selectedRequerimiento.Descripcion}</h2>
                <p>
                  Solicitado por {selectedRequerimiento.Solicitante?.Title || 'Sin solicitante'} · {' '}
                  {new Date(selectedRequerimiento.Created).toLocaleDateString('es-EC')}
                </p>
              </div>
              <div className={styles.detailAmount}>
                <small>Valor promedio de cotizaciones</small>
                <strong>{formatCurrency(selectedRequerimiento.ValorTotal)}</strong>
              </div>
            </div>

            <section className={styles.timelineCard}>
              <div className={styles.timelineTitle}>
                <div><span className={styles.eyebrow}>SEGUIMIENTO</span><h3>Etapas del proceso</h3></div>
                <strong>{getEtapaLabel(selectedRequerimiento.EtapaActual)}</strong>
              </div>
              <div className={styles.timeline}>
                {getEtapas(selectedRequerimiento).map((etapa, index, etapas) => {
                  const currentIndex = Math.max(0, etapas.indexOf(selectedRequerimiento.EtapaActual || 'Solicitud'));
                  const completed = index < currentIndex || selectedRequerimiento.EtapaActual === 'Finalizado';
                  const current = index === currentIndex && selectedRequerimiento.EtapaActual !== 'Finalizado';
                  return (
                    <div key={etapa} className={`${styles.timelineStep} ${completed ? styles.timelineCompleted : ''} ${current ? styles.timelineCurrent : ''}`}>
                      <div className={styles.timelineMarker}>{completed ? '✓' : index + 1}</div>
                      <span>{getEtapaLabel(etapa)}</span>
                    </div>
                  );
                })}
              </div>
            </section>

            <div className={styles.detailGrid}>
              <section className={styles.detailCard}>
                <span className={styles.eyebrow}>INFORMACIÓN GENERAL</span>
                <h3>Datos del requerimiento</h3>
                <div className={styles.infoGrid}>
                  <div><small>Categoría</small><strong>{getCategoriaNombre(selectedRequerimiento.CategoriaId)}</strong></div>
                  <div><small>Tipo de proceso</small><strong>{selectedRequerimiento.Recurrente ? 'Recurrente' : 'No recurrente'}</strong></div>
                  <div><small>Etapa actual</small><strong>{getEtapaLabel(selectedRequerimiento.EtapaActual)}</strong></div>
                  <div><small>Estado</small><strong>{selectedRequerimiento.Estado}</strong></div>
                </div>
                <div className={styles.descriptionBox}>
                  <small>Descripción</small>
                  <p>{selectedRequerimiento.Descripcion}</p>
                </div>
              </section>

              <aside className={styles.processCard}>
                <span className={styles.eyebrow}>SIGUIENTE PASO</span>
                <h3>{selectedRequerimiento.Estado === 'Borrador' ? 'Completar solicitud' : 'Seguimiento del expediente'}</h3>
                <p>
                  {selectedRequerimiento.Estado === 'Borrador'
                    ? 'El requerimiento permanece como borrador y todavía no ha ingresado al flujo de aprobación.'
                    : `El expediente se encuentra actualmente en ${getEtapaLabel(selectedRequerimiento.EtapaActual).toLowerCase()}.`}
                </p>
                            </aside>
            </div>

            {/* DETALLE DE COTIZACIONES */}

            <section className={styles.detailCard}>

              <span className={styles.eyebrow}>
                COTIZACIONES
              </span>

              <h3>Propuestas recibidas</h3>

              {loadingDetalle ? (

                <div className={styles.emptyState}>
                  Cargando cotizaciones...
                </div>

              ) : cotizacionesDetalle.length === 0 ? (

                <div className={styles.emptyState}>
                  Este requerimiento no tiene cotizaciones registradas.
                </div>

              ) : (

                <div className={styles.quotationList}>

                  {cotizacionesDetalle.map(
                    (cotizacion, cotizacionIndex) => (

                      <div
                        key={cotizacion.Id}
                        className={styles.quotationCard}
                      >

                        <div className={styles.quotationHeader}>

                          <div>

                            <span className={styles.quotationLabel}>
                              COTIZACIÓN {cotizacionIndex + 1}
                            </span>

                            <h3>
                              Propuesta comercial
                            </h3>

                            {cotizacion.CodigoCotizacionInterno && (
                              <small>
                                {cotizacion.CodigoCotizacionInterno}
                              </small>
                            )}

                          </div>

                          <div className={styles.quotationAmount}>
                            <small>Total cotización</small>
                            <strong>
                              {formatCurrency(cotizacion.ValorTotal)}
                            </strong>
                          </div>

                        </div>

                        {(user?.role === 'Aprobador' || user?.role === 'Admin') &&
                          selectedRequerimiento.Estado === 'Enviado Aprobacion' && (
                            <button
                              type="button"
                              className={`${styles.recurrentToggle} ${
                                cotizacionSeleccionadaId === cotizacion.Id
                                  ? styles.recurrentToggleActive
                                  : ''
                              }`}
                              style={{ width: '100%', marginTop: 16 }}
                              onClick={() => setCotizacionSeleccionadaId(cotizacion.Id)}
                              disabled={procesandoDecision}
                            >
                              <span className={styles.toggleDot} />
                              <span>
                                <strong>
                                  {cotizacionSeleccionadaId === cotizacion.Id
                                    ? 'Cotización seleccionada'
                                    : 'Seleccionar esta cotización'}
                                </strong>
                                <small>
                                  {cotizacionSeleccionadaId === cotizacion.Id
                                    ? 'Esta propuesta será aprobada'
                                    : 'Elige esta propuesta para continuar con la aprobación'}
                                </small>
                              </span>
                            </button>
                        )}

                        {cotizacion.Seleccionada &&
                          selectedRequerimiento.Estado === 'Aprobado' && (
                            <div style={{
                              marginTop: 16,
                              padding: '12px 16px',
                              borderRadius: 10,
                              background: '#f0fdf4',
                              border: '1px solid #bbf7d0'
                            }}>
                              <strong>✓ Cotización aprobada</strong>
                            </div>
                        )}

                        {cotizacion.Attachments.length > 0 && (
                          <div style={{ marginTop: 24 }}>
                            <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 12 }}>
                              <div>
                                <strong>Documento de cotización</strong>
                                <span>Vista previa del archivo adjunto</span>
                              </div>
                            </div>
                            {cotizacion.Attachments.map((archivo) => (
                              <div key={archivo.ServerRelativeUrl} style={{ border: '1px solid #e5e7eb', borderRadius: 12, overflow: 'hidden', background: '#fff', marginTop: 12 }}>
                                <div style={{ padding: '12px 16px', fontSize: 13, fontWeight: 600, borderBottom: '1px solid #e5e7eb', background: '#f8fafc' }}>{archivo.FileName}</div>
                                <iframe
                                  title={archivo.FileName}
                                  src={getVistaPreviaArchivo(archivo)}
                                  style={{ display: 'block', width: '100%', height: 650, border: 0, background: '#f3f4f6' }}
                                />
                              </div>
                            ))}
                          </div>
                        )}

                      </div>

                    )
                  )}

                </div>

              )}

            </section>

            {(user?.role === 'Aprobador' || user?.role === 'Admin') &&
              selectedRequerimiento.Estado === 'Enviado Aprobacion' && (
                <section className={styles.detailCard}>
                  <span className={styles.eyebrow}>DECISIÓN DEL APROBADOR</span>
                  <h3>Aprobar o rechazar requerimiento</h3>
                  <p>
                    Para aprobar, selecciona primero una de las cotizaciones mostradas arriba.
                    Al rechazar, debes ingresar un comentario.
                  </p>
                  <div className={styles.formGroup}>
                    <label className={styles.label}>Comentario (obligatorio al rechazar)</label>
                    <textarea
                      className={styles.textarea}
                      rows={4}
                      value={comentarioAprobador}
                      placeholder="Escribe una observación sobre la decisión..."
                      onChange={(event: React.ChangeEvent<HTMLTextAreaElement>) =>
                        setComentarioAprobador(event.target.value)
                      }
                    />
                  </div>
                  <div className={styles.actions}>
                    <button type="button" className={styles.secondaryButton} disabled={procesandoDecision}
                      onClick={() => { void decidirRequerimiento(false); }}>
                      {procesandoDecision ? 'Procesando...' : 'Rechazar'}
                    </button>
                    <button type="button" className={styles.primaryButton} disabled={procesandoDecision}
                      onClick={() => { void decidirRequerimiento(true); }}>
                      {procesandoDecision ? 'Procesando...' : 'Aprobar cotización'}
                    </button>
                  </div>
                </section>
              )}

          </div>
        ) : (
          <div className={styles.requestsWorkspace}>
            <div className={styles.pageIntro}>
              <div>
                <span className={styles.eyebrow}>CONTROL DE EXPEDIENTES</span>
                <h2 className={styles.pageTitle}>{activeView === 'mis' ? 'Mis solicitudes' : 'Solicitudes'}</h2>
                <p className={styles.pageDescription}>Consulta el avance, estado y etapa actual de cada requerimiento.</p>
              </div>
            </div>

            <div className={styles.kpiGrid}>
              <div className={styles.kpiCard}><span>Total</span><strong>{requerimientos.length}</strong><small>requerimientos</small></div>
              <div className={styles.kpiCard}><span>En aprobación</span><strong>{requerimientos.filter(r => r.Estado === 'Enviado Aprobacion').length}</strong><small>pendientes</small></div>
              <div className={styles.kpiCard}><span>Aprobados</span><strong>{requerimientos.filter(r => r.Estado === 'Aprobado').length}</strong><small>procesados</small></div>
              <div className={styles.kpiCard}><span>Borradores</span><strong>{requerimientos.filter(r => r.Estado === 'Borrador').length}</strong><small>sin enviar</small></div>
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
                  <thead><tr><th>ID</th><th>Requerimiento</th><th>Categoría</th><th>Promedio cotizaciones</th><th>Etapa actual</th><th>Estado</th><th>Fecha</th><th /></tr></thead>
                  <tbody>
                    {filteredRequerimientos.map((requerimiento: IRequerimientoItem) => (
                      <tr key={requerimiento.Id} className={styles.clickableRow} onClick={() => abrirRequerimiento(requerimiento)}>
                        <td><strong>#{requerimiento.Id}</strong></td>
                        <td><div className={styles.requestTitle}>{requerimiento.Descripcion}</div><small>{requerimiento.Solicitante?.Title || 'Sin solicitante'}</small></td>
                        <td>{getCategoriaNombre(requerimiento.CategoriaId)}</td>
                        <td><strong>{formatCurrency(requerimiento.ValorTotal)}</strong></td>
                        <td><span className={styles.stageBadge}>{getEtapaLabel(requerimiento.EtapaActual)}</span></td>
                        <td><span className={getEstadoClass(requerimiento.Estado)}>{requerimiento.Estado}</span></td>
                        <td>{new Date(requerimiento.Created).toLocaleDateString('es-EC')}</td>
                        <td><button type="button" className={styles.viewButton} onClick={(e) => { e.stopPropagation(); abrirRequerimiento(requerimiento); }}>Ver →</button></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          </div>
        )
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
