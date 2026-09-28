import * as React from 'react';
import './global/global.module.scss';
import { useEffect, useRef, useState } from 'react';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import styles from './Requerimientos.module.scss';
import inovaLogo from '../assets/inova-logo.png';
import { IRequerimientosProps } from './IRequerimientosProps';

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

  const [archivo, setArchivo] =
    useState<File | null>(null);

  const fileInputRef =
    useRef<HTMLInputElement>(null);

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
        `?$select=Id,NombreCategoria,AreaId,Activo` +
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
        (item: { Id: number; NombreCategoria?: string }) => ({
          Id: item.Id,
          NombreCategoria: item.NombreCategoria || `Categoría ${item.Id}`
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
              `Estado eq 'Enviado Aprobacion'`;
            break;

          case 'aprobadas':
            filter =
              `Estado eq 'Aprobado'`;
            break;

          case 'rechazadas':
            filter =
              `Estado eq 'Rechazado'`;
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
          `Recurrente,Created,` +
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

  /*
   * ===========================
   * CREAR REQUERIMIENTO
   * DESTINO: LISTA SHAREPOINT
   * ===========================
   */

  const createRequerimiento =
    async (): Promise<void> => {

      try {

        setSaving(true);
        setMessage('');

        if (!user) {
          throw new Error(
            'No se pudo obtener el usuario actual.'
          );
        }

        if (
          user.role !== 'Solicitador' &&
          user.role !== 'Admin'
        ) {
          throw new Error(
            'No tienes permisos para crear requerimientos.'
          );
        }

        if (!descripcion.trim()) {
          throw new Error(
            'Ingrese una descripción.'
          );
        }

        if (!categoriaId) {
          throw new Error(
            'Seleccione una categoría.'
          );
        }

        if (!valorTotal) {
          throw new Error(
            'Ingrese el valor total.'
          );
        }

        const numericValue =
          Number(valorTotal);

        if (
          isNaN(numericValue) ||
          numericValue < 0
        ) {
          throw new Error(
            'Ingrese un valor total válido.'
          );
        }

        if (!archivo) {
          throw new Error(
            'Seleccione un archivo PDF.'
          );
        }

        const isPdf =
          archivo.type === 'application/pdf' ||
          /\.pdf$/i.test(archivo.name);

        if (!isPdf) {
          throw new Error(
            'El archivo debe ser un PDF.'
          );
        }

        /*
         * Crear elemento en Requerimientos
         */

        const createUrl =
          `${context.pageContext.web.absoluteUrl}` +
          `/_api/web/lists/getbytitle('Requerimientos')/items`;

        const body = {

          Title:
            descripcion
              .trim()
              .substring(0, 255),

          Descripcion:
            descripcion.trim(),

          SolicitanteId:
            user.id,

          CategoriaId:
            Number(categoriaId),

          Recurrente:
            recurrente,

          ValorTotal:
            numericValue,

          Estado:
            'Enviado Aprobacion'
        };

        const createResponse:
          SPHttpClientResponse =
            await context.spHttpClient.post(
              createUrl,
              SPHttpClient.configurations.v1,
              {
                headers: {
                  Accept:
                    'application/json;odata=nometadata',

                  'Content-Type':
                    'application/json;odata=nometadata'
                },

                body:
                  JSON.stringify(body)
              }
            );

        if (!createResponse.ok) {

          const responseText =
            await createResponse.text();

          throw new Error(
            `Error creando requerimiento. ` +
            `HTTP ${createResponse.status}: ${responseText}`
          );
        }

        const createdItem =
          await createResponse.json();

        const itemId: number =
          createdItem.Id ||
          createdItem.ID;

        if (!itemId) {
          throw new Error(
            'El requerimiento se creó, pero no se pudo obtener su ID.'
          );
        }

        /*
         * Adjuntar PDF al elemento creado
         */

        const safeFileName =
          archivo.name.replace(
            /'/g,
            "''"
          );

        const attachmentUrl =
          `${context.pageContext.web.absoluteUrl}` +
          `/_api/web/lists/getbytitle('Requerimientos')` +
          `/items(${itemId})` +
          `/AttachmentFiles/add(FileName=@fileName)` +
          `?@fileName='${encodeURIComponent(safeFileName)}'`;

        const attachmentResponse:
          SPHttpClientResponse =
            await context.spHttpClient.post(
              attachmentUrl,
              SPHttpClient.configurations.v1,
              {
                headers: {
                  Accept:
                    'application/json;odata=nometadata'
                },

                body:
                  archivo
              }
            );

        if (!attachmentResponse.ok) {

          const attachmentError =
            await attachmentResponse.text();

          throw new Error(
            `El requerimiento #${itemId} fue creado, ` +
            `pero el PDF no pudo adjuntarse. ` +
            `HTTP ${attachmentResponse.status}: ${attachmentError}`
          );
        }

        /*
         * Limpiar formulario
         */

        setDescripcion('');
        setCategoriaId('');
        setRecurrente(false);
        setValorTotal('');
        setArchivo(null);

        if (
          fileInputRef.current
        ) {
          fileInputRef.current.value = '';
        }

        /*
         * Después de crear:
         * Solicitador -> Mis solicitudes
         * Admin -> Todas
         */

        if (
          user.role === 'Solicitador'
        ) {

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
          `Requerimiento #${itemId} creado correctamente con su PDF adjunto.`
        );

      } catch (error) {

        setMessage(
          getErrorMessage(error)
        );

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
                  {user.role}
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

      {/* DATOS DEL USUARIO */}

      {user && (

        <div
          className={
            styles.userCard
          }
        >

          <div
            className={
              styles.userMainInfo
            }
          >

            <div
              className={
                styles.avatarWrapper
              }
            >

              <img
                className={
                  styles.avatar
                }
                src={
                  `${context.pageContext.web.absoluteUrl}` +
                  `/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(user.email)}`
                }
                alt={
                  user.displayName
                }
              />

              <span
                className={
                  styles.onlineIndicator
                }
              />

            </div>

            <div>

              <div
                className={
                  styles.userName
                }
              >
                {user.displayName}
              </div>

              <div
                className={
                  styles.secondaryText
                }
              >
                {user.email}
              </div>

              <div
                className={
                  styles.userMeta
                }
              >
                <span>
                  Rol:
                </span>

                <strong>
                  {user.role}
                </strong>
              </div>

              {user.area && (
                <div className={styles.userMeta}>
                  <span>Área:</span>
                  <strong>{user.area}</strong>
                </div>
              )}

            </div>

          </div>

          <div
            className={
              styles.userMessage
            }
          >
            <span
              className={
                styles.quote
              }
            >
              “
            </span>

            <span>
              Juntos hacemos posibles
              <br />
              las grandes soluciones.
            </span>
          </div>

        </div>
      )}

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
      ) &&
      activeView === 'nueva' && (

        <div className={styles.card}>

          <div
            className={
              styles.formGroup
            }
          >

            <label
              className={
                styles.label
              }
            >
              Descripción
            </label>

            <textarea
              className={
                styles.textarea
              }
              value={
                descripcion
              }
              placeholder="Describe brevemente tu requerimiento..."
              onChange={(
                event:
                React.ChangeEvent<HTMLTextAreaElement>
              ) =>
                setDescripcion(
                  event.target.value
                )
              }
              rows={5}
            />

          </div>

          <div
            className={
              styles.formGroup
            }
          >

            <label
              className={
                styles.label
              }
            >
              Categoría
            </label>

            <select
              className={
                styles.select
              }
              value={
                categoriaId
              }
              onChange={(
                event:
                React.ChangeEvent<HTMLSelectElement>
              ) =>
                setCategoriaId(
                  event.target.value
                )
              }
            >

              <option value="">
                Seleccione una categoría
              </option>

              {categories.map(
                (
                  category:
                  ICategory
                ) => (

                  <option
                    key={
                      category.Id
                    }
                    value={
                      category.Id
                    }
                  >
                    {
                      category.NombreCategoria
                    }
                  </option>
                )
              )}

            </select>

          </div>

          <div
            className={
              styles.formGroup
            }
          >

            <label
              className={
                styles.checkboxLabel
              }
            >

              <input
                type="checkbox"
                checked={
                  recurrente
                }
                onChange={(
                  event:
                  React.ChangeEvent<HTMLInputElement>
                ) =>
                  setRecurrente(
                    event.target.checked
                  )
                }
              />

              <span>
                Requerimiento recurrente
              </span>

            </label>

          </div>

          <div
            className={
              styles.formGroup
            }
          >

            <label
              className={
                styles.label
              }
            >
              Valor total
            </label>

            <input
              className={
                styles.input
              }
              type="number"
              min="0"
              step="0.01"
              placeholder="0.00"
              value={
                valorTotal
              }
              onChange={(
                event:
                React.ChangeEvent<HTMLInputElement>
              ) =>
                setValorTotal(
                  event.target.value
                )
              }
            />

          </div>

          <div
            className={
              styles.fileBox
            }
          >

            <label
              className={
                styles.label
              }
            >
              Documento PDF
            </label>

            <input
              ref={
                fileInputRef
              }
              type="file"
              accept="application/pdf,.pdf"
              onChange={(
                event:
                React.ChangeEvent<HTMLInputElement>
              ) => {

                const selectedFile =
                  event.target.files?.[0];

                if (!selectedFile) {

                  setArchivo(null);
                  return;
                }

                const isPdf =
                  selectedFile.type ===
                    'application/pdf' ||
                  /\.pdf$/i.test(
                    selectedFile.name
                  );

                if (!isPdf) {

                  setArchivo(null);

                  setMessage(
                    'Solo se permiten archivos PDF.'
                  );

                  event.target.value = '';

                  return;
                }

                setMessage('');
                setArchivo(
                  selectedFile
                );
              }}
            />

            {archivo && (

              <div
                className={
                  styles.fileName
                }
              >
                Archivo seleccionado:{' '}
                <strong>
                  {archivo.name}
                </strong>
              </div>
            )}

          </div>

          <div
            className={
              styles.actions
            }
          >

            <button
              type="button"
              className={
                styles.primaryButton
              }
              disabled={
                saving
              }
              onClick={() => {
                void createRequerimiento();
              }}
            >
              {
                saving
                  ? 'Enviando...'
                  : 'Enviar requerimiento'
              }
            </button>

          </div>

        </div>
      )}

      {/* TABLA DE SOLICITUDES */}

      {activeView !== 'nueva' &&
       user?.role !== 'SinRol' && (

        <div
          className={
            styles.requestsSection
          }
        >

          {loadingRequerimientos ? (

            <div
              className={
                styles.emptyState
              }
            >
              Cargando solicitudes...
            </div>

          ) : requerimientos.length === 0 ? (

            <div
              className={
                styles.emptyState
              }
            >
              No existen solicitudes
              para mostrar.
            </div>

          ) : (

            <table
              className={
                styles.requestsTable
              }
            >

              <thead>
                <tr>
                  <th>ID</th>
                  <th>Solicitante</th>
                  <th>Descripción</th>
                  <th>Categoría</th>
                  <th>Recurrente</th>
                  <th>Valor</th>
                  <th>Estado</th>
                  <th>Fecha</th>
                </tr>
              </thead>

              <tbody>

                {requerimientos.map(
                  (
                    requerimiento:
                    IRequerimientoItem
                  ) => (

                    <tr
                      key={
                        requerimiento.Id
                      }
                    >

                      <td>
                        #{requerimiento.Id}
                      </td>

                      <td>
                        {
                          requerimiento
                            .Solicitante
                            ?.Title ||
                          'Sin solicitante'
                        }
                      </td>

                      <td>
                        {
                          requerimiento
                            .Descripcion
                        }
                      </td>

                      <td>
                        {
                          getCategoriaNombre(
                            requerimiento
                              .CategoriaId
                          )
                        }
                      </td>

                      <td>
                        {
                          requerimiento
                            .Recurrente
                            ? 'Sí'
                            : 'No'
                        }
                      </td>

                      <td>
                        {
                          Number(
                            requerimiento
                              .ValorTotal
                          ).toLocaleString(
                            'es-EC',
                            {
                              style:
                                'currency',
                              currency:
                                'USD'
                            }
                          )
                        }
                      </td>

                      <td>
                        {
                          requerimiento
                            .Estado
                        }
                      </td>

                      <td>
                        {
                          new Date(
                            requerimiento
                              .Created
                          ).toLocaleDateString(
                            'es-EC'
                          )
                        }
                      </td>

                    </tr>
                  )
                )}

              </tbody>

            </table>
          )}

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
