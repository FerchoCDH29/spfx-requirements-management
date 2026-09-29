import * as React from 'react';
import './global/global.module.scss';
import { useEffect, useState } from 'react';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

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

interface IProveedor {
  Id: number;
  RazonSocial: string;
  RUC?: string;
}

interface IItemCotizacionForm {
  key: string;
  descripcion: string;
  cantidad: string;
  valorUnitario: string;
}

interface ICotizacionForm {
  key: string;
  proveedorId: string;
  descripcion: string;
  fechaCotizacion: string;
  archivo: File | null;
  items: IItemCotizacionForm[];
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

  const [valorTotal, setValorTotal] =
    useState<string>('');

  const [proveedores, setProveedores] =
    useState<IProveedor[]>([]);

  const createEmptyItem = (): IItemCotizacionForm => ({
    key: `${Date.now()}-${Math.random()}`,
    descripcion: '',
    cantidad: '1',
    valorUnitario: ''
  });

  const createEmptyCotizacion = (): ICotizacionForm => ({
    key: `${Date.now()}-${Math.random()}`,
    proveedorId: '',
    descripcion: '',
    fechaCotizacion: '',
    archivo: null,
    items: [createEmptyItem()]
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
   * PROVEEDORES
   * ===========================
   */

  const loadProveedores = async (): Promise<void> => {
    const url =
      `${context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists/getbytitle('Proveedores')/items` +
      `?$select=Id,RazonSocial,RUC,Activo` +
      `&$filter=${encodeURIComponent('Activo eq 1')}` +
      `&$orderby=RazonSocial asc&$top=500`;

    const response = await context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=nometadata' } }
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(
        `Error cargando proveedores. HTTP ${response.status}: ${errorText}`
      );
    }

    const data = await response.json();
    setProveedores(
      data.value.map((item: { Id: number; RazonSocial?: string; RUC?: string }) => ({
        Id: item.Id,
        RazonSocial: item.RazonSocial || `Proveedor ${item.Id}`,
        RUC: item.RUC
      }))
    );
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

          await loadProveedores();

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

  /*
   * ===========================
   * GUARDAR / ENVIAR REQUERIMIENTO
   * ===========================
   */

  const calcularTotalCotizacion = (cotizacion: ICotizacionForm): number =>
    cotizacion.items.reduce((total, item) => {
      const cantidad = Number(item.cantidad) || 0;
      const valorUnitario = Number(item.valorUnitario) || 0;
      return total + (cantidad * valorUnitario);
    }, 0);

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

  const actualizarCotizacion = (
    cotizacionIndex: number,
    changes: Partial<ICotizacionForm>
  ): void => {
    setCotizaciones(
      cotizaciones.map((cotizacion, index) =>
        index === cotizacionIndex ? { ...cotizacion, ...changes } : cotizacion
      )
    );
  };

  const agregarItem = (cotizacionIndex: number): void => {
    setCotizaciones(
      cotizaciones.map((cotizacion, index) =>
        index === cotizacionIndex
          ? { ...cotizacion, items: [...cotizacion.items, createEmptyItem()] }
          : cotizacion
      )
    );
  };

  const eliminarItem = (cotizacionIndex: number, itemIndex: number): void => {
    const cotizacion = cotizaciones[cotizacionIndex];
    if (cotizacion.items.length === 1) {
      setMessage('Cada cotización debe mantener al menos un ítem.');
      return;
    }

    setCotizaciones(
      cotizaciones.map((current, index) =>
        index === cotizacionIndex
          ? { ...current, items: current.items.filter((_, i) => i !== itemIndex) }
          : current
      )
    );
  };

  const actualizarItem = (
    cotizacionIndex: number,
    itemIndex: number,
    changes: Partial<IItemCotizacionForm>
  ): void => {
    setCotizaciones(
      cotizaciones.map((cotizacion, index) =>
        index === cotizacionIndex
          ? {
              ...cotizacion,
              items: cotizacion.items.map((item, i) =>
                i === itemIndex ? { ...item, ...changes } : item
              )
            }
          : cotizacion
      )
    );
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

      const numericValue = Number(valorTotal);
      if (!valorTotal || isNaN(numericValue) || numericValue < 0) {
        throw new Error('Ingrese un valor total válido.');
      }

      if (cotizaciones.length < 1 || cotizaciones.length > 3) {
        throw new Error('El requerimiento debe tener entre 1 y 3 cotizaciones.');
      }

      if (enviar) {
        cotizaciones.forEach((cotizacion, cotizacionIndex) => {
          if (!cotizacion.proveedorId) {
            throw new Error(`Seleccione el proveedor de la cotización ${cotizacionIndex + 1}.`);
          }
          if (!cotizacion.descripcion.trim()) {
            throw new Error(`Ingrese la descripción de la cotización ${cotizacionIndex + 1}.`);
          }
          if (cotizacion.items.length < 1) {
            throw new Error(`La cotización ${cotizacionIndex + 1} debe tener al menos un ítem.`);
          }
          cotizacion.items.forEach((item, itemIndex) => {
            if (!item.descripcion.trim()) {
              throw new Error(
                `Ingrese la descripción del ítem ${itemIndex + 1} de la cotización ${cotizacionIndex + 1}.`
              );
            }
            const cantidad = Number(item.cantidad);
            const valorUnitarioItem = Number(item.valorUnitario);
            if (!cantidad || cantidad <= 0) {
              throw new Error(
                `Ingrese una cantidad válida en el ítem ${itemIndex + 1} de la cotización ${cotizacionIndex + 1}.`
              );
            }
            if (isNaN(valorUnitarioItem) || valorUnitarioItem < 0 || item.valorUnitario === '') {
              throw new Error(
                `Ingrese un valor unitario válido en el ítem ${itemIndex + 1} de la cotización ${cotizacionIndex + 1}.`
              );
            }
          });
        });
      }

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
          !!cotizacion.proveedorId ||
          !!cotizacion.descripcion.trim() ||
          !!cotizacion.fechaCotizacion ||
          !!cotizacion.archivo ||
          cotizacion.items.some(item =>
            !!item.descripcion.trim() || !!item.valorUnitario
          );

        if (!enviar && !tieneDatos) {
          continue;
        }

        const totalCotizacion = calcularTotalCotizacion(cotizacion);
        const cotizacionBody: Record<string, unknown> = {
          Title: `REQ-${itemId}-COT-${c + 1}`,
          RequerimientoId: itemId,
          Descripcion: cotizacion.descripcion.trim(),
          CodigoCotizacionInterno: `COT-REQ-${('000000' + itemId).slice(-6)}-${('00' + (c + 1)).slice(-2)}`,
          ValorTotal: totalCotizacion,
          Seleccionada: false,
          Activo: true
        };

        if (cotizacion.proveedorId) {
          cotizacionBody.ProveedorId = Number(cotizacion.proveedorId);
        }
        if (cotizacion.fechaCotizacion) {
          cotizacionBody.FechaCotizacion = cotizacion.fechaCotizacion;
        }

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

        for (let i = 0; i < cotizacion.items.length; i += 1) {
          const item = cotizacion.items[i];
          const tieneDatosItem = !!item.descripcion.trim() || !!item.valorUnitario;
          if (!enviar && !tieneDatosItem) {
            continue;
          }

          const cantidad = Number(item.cantidad) || 0;
          const valorUnitarioItem = Number(item.valorUnitario) || 0;

          await createListItem('ItemCotizaciones', {
            Title: `COT-${cotizacionId}-ITEM-${i + 1}`,
            CotizacionId: cotizacionId,
            Descripcion: item.descripcion.trim(),
            Cantidad: cantidad,
            ValorUnitario: valorUnitarioItem,
            ValorTotal: cantidad * valorUnitarioItem,
            Seleccionado: false
          });
        }
      }

      setDescripcion('');
      setCategoriaId('');
      setRecurrente(false);
      setValorTotal('');
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
                Registra la solicitud y compara las cotizaciones antes de enviarla a aprobación.
              </p>
            </div>
            <div className={styles.formProgress}>Solicitud · Paso inicial</div>
          </div>

          <section className={styles.formSection}>
            <div className={styles.sectionHeader}>
              <div className={styles.sectionNumber}>01</div>
              <div>
                <h3>Información general</h3>
                <p>Datos principales que identificarán el expediente.</p>
              </div>
            </div>

            <div className={styles.formGroup}>
              <label className={styles.label}>Descripción del requerimiento</label>
              <textarea
                className={styles.textarea}
                value={descripcion}
                placeholder="Describe el producto, servicio o necesidad..."
                onChange={(event: React.ChangeEvent<HTMLTextAreaElement>) =>
                  setDescripcion(event.target.value)
                }
                rows={4}
              />
            </div>

            <div className={styles.formGrid3}>
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
                <label className={styles.label}>Valor referencial</label>
                <div className={styles.moneyInput}>
                  <span>$</span>
                  <input
                    className={styles.input}
                    type="number"
                    min="0"
                    step="0.01"
                    placeholder="0.00"
                    value={valorTotal}
                    onChange={(event: React.ChangeEvent<HTMLInputElement>) =>
                      setValorTotal(event.target.value)
                    }
                  />
                </div>
              </div>

              <div className={styles.formGroup}>
                <label className={styles.label}>Tipo de proceso</label>
                <button
                  type="button"
                  className={`${styles.recurrentToggle} ${recurrente ? styles.recurrentToggleActive : ''}`}
                  onClick={() => setRecurrente(!recurrente)}
                >
                  <span className={styles.toggleDot} />
                  <span>
                    <strong>{recurrente ? 'Recurrente' : 'No recurrente'}</strong>
                    <small>{recurrente ? 'Incluye etapa de contrato' : 'Pasa directo a orden de compra'}</small>
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
                  <p>Compara hasta tres propuestas. Cada una puede incluir múltiples ítems.</p>
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
                      <h3>{cotizacion.proveedorId
                        ? proveedores.find(p => p.Id === Number(cotizacion.proveedorId))?.RazonSocial || 'Proveedor'
                        : 'Nueva propuesta'}</h3>
                    </div>
                    <div className={styles.quotationAmount}>
                      <small>Total cotización</small>
                      <strong>{formatCurrency(calcularTotalCotizacion(cotizacion))}</strong>
                    </div>
                  </div>

                  <div className={styles.quotationMetaGrid}>
                    <div className={styles.formGroup}>
                      <label className={styles.label}>Proveedor</label>
                      <select
                        className={styles.select}
                        value={cotizacion.proveedorId}
                        onChange={(event: React.ChangeEvent<HTMLSelectElement>) =>
                          actualizarCotizacion(cotizacionIndex, { proveedorId: event.target.value })
                        }
                      >
                        <option value="">Seleccione un proveedor</option>
                        {proveedores.map((proveedor: IProveedor) => (
                          <option key={proveedor.Id} value={proveedor.Id}>{proveedor.RazonSocial}</option>
                        ))}
                      </select>
                    </div>

                    <div className={styles.formGroup}>
                      <label className={styles.label}>Fecha</label>
                      <input
                        className={styles.input}
                        type="date"
                        value={cotizacion.fechaCotizacion}
                        onChange={(event: React.ChangeEvent<HTMLInputElement>) =>
                          actualizarCotizacion(cotizacionIndex, { fechaCotizacion: event.target.value })
                        }
                      />
                    </div>

                    <div className={styles.formGroup}>
                      <label className={styles.label}>Documento PDF (opcional)</label>
                      <label className={styles.fileUpload}>
                        <input
                          type="file"
                          accept="application/pdf,.pdf"
                          onChange={(event: React.ChangeEvent<HTMLInputElement>) => {
                            const selectedFile = event.target.files?.[0] || null;
                            if (selectedFile) {
                              const isPdf = selectedFile.type === 'application/pdf' || /\.pdf$/i.test(selectedFile.name);
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
                        <span>{cotizacion.archivo ? '✓ PDF adjunto' : 'Adjuntar PDF'}</span>
                      </label>
                    </div>
                  </div>

                  <div className={styles.formGroup}>
                    <label className={styles.label}>Descripción de la propuesta</label>
                    <input
                      className={styles.input}
                      type="text"
                      placeholder="Ej. Licenciamiento anual y soporte"
                      value={cotizacion.descripcion}
                      onChange={(event: React.ChangeEvent<HTMLInputElement>) =>
                        actualizarCotizacion(cotizacionIndex, { descripcion: event.target.value })
                      }
                    />
                  </div>

                  <div className={styles.itemsSection}>
                    <div className={styles.itemsHeader}>
                      <div>
                        <strong>Detalle de ítems</strong>
                        <span>{cotizacion.items.length} {cotizacion.items.length === 1 ? 'ítem' : 'ítems'}</span>
                      </div>
                      <button type="button" className={styles.secondaryButton} onClick={() => agregarItem(cotizacionIndex)}>
                        + Agregar ítem
                      </button>
                    </div>

                    <div className={styles.itemsTableHeader}>
                      <span>Descripción</span><span>Cantidad</span><span>Valor unitario</span><span>Total</span><span />
                    </div>

                    {cotizacion.items.map((item, itemIndex) => {
                      const totalItem = (Number(item.cantidad) || 0) * (Number(item.valorUnitario) || 0);
                      return (
                        <div key={item.key} className={styles.itemRow}>
                          <input
                            className={styles.input}
                            type="text"
                            placeholder="Producto o servicio"
                            value={item.descripcion}
                            onChange={(event: React.ChangeEvent<HTMLInputElement>) =>
                              actualizarItem(cotizacionIndex, itemIndex, { descripcion: event.target.value })
                            }
                          />
                          <input
                            className={styles.input}
                            type="number"
                            min="1"
                            step="1"
                            value={item.cantidad}
                            onChange={(event: React.ChangeEvent<HTMLInputElement>) =>
                              actualizarItem(cotizacionIndex, itemIndex, { cantidad: event.target.value })
                            }
                          />
                          <input
                            className={styles.input}
                            type="number"
                            min="0"
                            step="0.01"
                            placeholder="0.00"
                            value={item.valorUnitario}
                            onChange={(event: React.ChangeEvent<HTMLInputElement>) =>
                              actualizarItem(cotizacionIndex, itemIndex, { valorUnitario: event.target.value })
                            }
                          />
                          <strong className={styles.itemTotal}>{formatCurrency(totalItem)}</strong>
                          <button
                            type="button"
                            className={styles.iconButton}
                            disabled={cotizacion.items.length === 1}
                            onClick={() => eliminarItem(cotizacionIndex, itemIndex)}
                            title="Eliminar ítem"
                          >×</button>
                        </div>
                      );
                    })}
                  </div>

                  <div className={styles.quotationFooter}>
                    {cotizaciones.length > 1 && (
                      <button type="button" className={styles.dangerLink} onClick={() => eliminarCotizacion(cotizacionIndex)}>
                        Eliminar cotización
                      </button>
                    )}
                    <div className={styles.quotationGrandTotal}>
                      <span>Total</span>
                      <strong>{formatCurrency(calcularTotalCotizacion(cotizacion))}</strong>
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
                <small>Valor referencial</small>
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
                  <div><small>Tipo</small><strong>{selectedRequerimiento.Recurrente ? 'Recurrente' : 'No recurrente'}</strong></div>
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
                  <thead><tr><th>ID</th><th>Requerimiento</th><th>Categoría</th><th>Monto</th><th>Etapa actual</th><th>Estado</th><th>Fecha</th><th /></tr></thead>
                  <tbody>
                    {filteredRequerimientos.map((requerimiento: IRequerimientoItem) => (
                      <tr key={requerimiento.Id} className={styles.clickableRow} onClick={() => setSelectedRequerimiento(requerimiento)}>
                        <td><strong>#{requerimiento.Id}</strong></td>
                        <td><div className={styles.requestTitle}>{requerimiento.Descripcion}</div><small>{requerimiento.Solicitante?.Title || 'Sin solicitante'}</small></td>
                        <td>{getCategoriaNombre(requerimiento.CategoriaId)}</td>
                        <td><strong>{formatCurrency(requerimiento.ValorTotal)}</strong></td>
                        <td><span className={styles.stageBadge}>{getEtapaLabel(requerimiento.EtapaActual)}</span></td>
                        <td><span className={getEstadoClass(requerimiento.Estado)}>{requerimiento.Estado}</span></td>
                        <td>{new Date(requerimiento.Created).toLocaleDateString('es-EC')}</td>
                        <td><button type="button" className={styles.viewButton} onClick={(e) => { e.stopPropagation(); setSelectedRequerimiento(requerimiento); }}>Ver →</button></td>
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
