import * as React from 'react';
import { useEffect, useRef, useState } from 'react';
import styles from './Requerimientos.module.scss';
import inovaLogo from '../assets/inova-logo.png';

import {
  MSGraphClientV3,
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import { IRequerimientosProps } from './IRequerimientosProps';

interface IUser {
  displayName: string;
  email: string;
  department: string;
}

interface IGraphUser {
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
  department?: string;
}

interface ICategory {
  Id: number;
  NombreCategoria: string;
}

interface ISharePointUser {
  Id: number;
  Title: string;
  Email: string;
}

const Requerimientos: React.FC<IRequerimientosProps> = ({
  context
}) => {

  const [user, setUser] = useState<IUser | null>(null);

  const [categories, setCategories] =
    useState<ICategory[]>([]);

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

  const [loading, setLoading] =
    useState<boolean>(true);

  const [saving, setSaving] =
    useState<boolean>(false);

  const [message, setMessage] =
    useState<string>('');

  const getErrorMessage = (error: unknown): string => {
    if (error instanceof Error) {
      return error.message;
    }

    return 'Ocurrió un error inesperado.';
  };

  useEffect(() => {

    const loadInitialData = async (): Promise<void> => {

      try {

        /*
         * ===========================
         * USUARIO MICROSOFT 365
         * ===========================
         */

        const graphClient: MSGraphClientV3 =
          await context.msGraphClientFactory.getClient('3');

        const graphUser: IGraphUser =
          await graphClient
            .api('/me')
            .select(
              'displayName,mail,userPrincipalName,department'
            )
            .get();

        setUser({
          displayName:
            graphUser.displayName || 'Usuario',

          email:
            graphUser.mail ||
            graphUser.userPrincipalName ||
            '',

          department:
            graphUser.department || ''
        });

        /*
         * ===========================
         * CATEGORIAS SHAREPOINT
         * ===========================
         */

        const categoriesUrl =
        `${context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('Categorias')/items?$top=100`;

      const categoriesResponse: SPHttpClientResponse =
        await context.spHttpClient.get(
          categoriesUrl,
          SPHttpClient.configurations.v1
        );

      if (!categoriesResponse.ok) {
        const errorText = await categoriesResponse.text();

        throw new Error(
          `Error cargando categorías. HTTP ${categoriesResponse.status}: ${errorText}`
        );
      }

      const categoriesData = await categoriesResponse.json();

      console.log(
        'CATEGORIAS SHAREPOINT:',
        categoriesData.value
      );

      const mappedCategories: ICategory[] =
        categoriesData.value.map((item: any) => ({
          Id: item.Id,
          NombreCategoria:
            item.NombreCategoria ||
            item.Title ||
            `Categoría ${item.Id}`
        }));

      setCategories(mappedCategories);

      } catch (error: unknown) {

        setMessage(getErrorMessage(error));

      } finally {

        setLoading(false);

      }
    };

    void loadInitialData();

  }, [context]);



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

        if (!archivo) {
        throw new Error(
          'Seleccione un archivo PDF.'
        );
      }

      if (archivo.type !== 'application/pdf') {
        throw new Error(
          'El archivo debe ser un PDF.'
        );
      }

        /*
         * ===========================
         * ID USUARIO SHAREPOINT
         * ===========================
         */

        const currentUserUrl =
          `${context.pageContext.web.absoluteUrl}` +
          `/_api/web/currentuser` +
          `?$select=Id,Title,Email`;

        const currentUserResponse:
          SPHttpClientResponse =
            await context.spHttpClient.get(
              currentUserUrl,
              SPHttpClient.configurations.v1
            );

        if (!currentUserResponse.ok) {
          throw new Error(
            'No se pudo obtener el usuario de SharePoint.'
          );
        }

        const currentUser:
          ISharePointUser =
            await currentUserResponse.json();

        /*
         * ===========================
         * CREAR REQUERIMIENTO
         * ===========================
         */

        const createUrl =
          `${context.pageContext.web.absoluteUrl}` +
          `/_api/web/lists/getbytitle('Requerimientos')/items`;

        const body = {

          /*
           * Lo ponemos también en Title para evitar
           * problemas si la columna Title sigue
           * siendo obligatoria.
           */
          Title: descripcion
            .trim()
            .substring(0, 255),

          Descripcion:
            descripcion.trim(),

          /*
           * Persona / Grupo
           */
          SolicitanteId:
            currentUser.Id,

          /*
           * Por ahora Area es texto.
           */
          Area:
            user.department ||
            'Sin departamento',

          /*
           * Lookup a Categorias
           */
          CategoriaId:
            Number(categoriaId),

          Recurrente:
            recurrente,

          ValorTotal:
            Number(valorTotal),

          Estado:
            'Enviado Aprobación'
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

                body: JSON.stringify(body)
              }
            );

        if (!createResponse.ok) {

            const responseText =
              await createResponse.text();

            throw new Error(
              `Error creando requerimiento. HTTP ${createResponse.status}: ${responseText}`
            );
          }

          /*
          * ===========================
          * OBTENER ID DEL REQUERIMIENTO
          * ===========================
          */

          const createdItem =
            await createResponse.json();

          const itemId: number =
            createdItem.Id || createdItem.ID;

          if (!itemId) {
            throw new Error(
              'El requerimiento se creó, pero no se pudo obtener su ID.'
            );
          }

          /*
          * ===========================
          * SUBIR PDF COMO ADJUNTO
          * ===========================
          */

          if (archivo) {

            const safeFileName =
              archivo.name.replace(/'/g, "''");

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

                    body: archivo
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
          }

          setMessage(
            `Requerimiento #${itemId} creado correctamente con su PDF adjunto.`
          );

        /*
         * Limpiar formulario
         */

        setDescripcion('');
        setCategoriaId('');
        setRecurrente(false);
        setValorTotal('');
        setArchivo(null);

        if (fileInputRef.current) {
          fileInputRef.current.value = '';
        }

      } catch (error: unknown) {

        setMessage(getErrorMessage(error));

      } finally {

        setSaving(false);

      }
    };



  if (loading) {
    return (
      <div>
        Cargando información...
      </div>
    );
  }



  return (

  <div className={styles.requerimientos}>

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
          <div className={styles.headerUserText}>
            <strong>
              {user.displayName}
            </strong>

            <span>
              {user.department || 'Sin departamento'}
            </span>
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


  <div className={styles.header}>
    <span className={styles.sectionLabel}>
      SOLICITUDES
    </span>

    <h1 className={styles.title}>
      Gestión de Requerimientos
    </h1>

    <p className={styles.subtitle}>
      Registra una nueva solicitud para iniciar el proceso de aprobación.
    </p>
  </div>
      {user && (
          <div className={styles.userCard}>

            <div className={styles.userMainInfo}>

              <div className={styles.avatarWrapper}>
                <img
                  className={styles.avatar}
                  src={
                    `${context.pageContext.web.absoluteUrl}` +
                    `/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(user.email)}`
                  }
                  alt={user.displayName}
                />

                <span className={styles.onlineIndicator} />
              </div>


              <div>
                <div className={styles.userName}>
                  {user.displayName}
                </div>

                <div className={styles.secondaryText}>
                  {user.email}
                </div>

                <div className={styles.userMeta}>
                  <span>
                    Área:
                  </span>

                  <strong>
                    {user.department || 'Sin departamento'}
                  </strong>
                </div>
              </div>

            </div>


            <div className={styles.userMessage}>
              <span className={styles.quote}>
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


     <div className={styles.formGroup}>

      <label className={styles.label}>
        Descripción
      </label>

      <textarea
        className={styles.textarea}
        value={descripcion}
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

      <div className={styles.formGroup}>

        <label className={styles.label}>
          Categoría
        </label>

        <select
          className={styles.select}
          value={categoriaId}
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
            (category: ICategory) => (

              <option
                key={category.Id}
                value={category.Id}
              >
                {category.NombreCategoria}
              </option>

            )
          )}

        </select>

      </div>

   

      <div className={styles.formGroup}>

        <label className={styles.checkboxLabel}>

          <input
            type="checkbox"
            checked={recurrente}
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


      <div className={styles.formGroup}>

        <label className={styles.label}>
          Valor total
        </label>

        <input
          className={styles.input}
          type="number"
          min="0"
          step="0.01"
          placeholder="0.00"
          value={valorTotal}
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
        style={{
          marginBottom: '18px'
        }}
      >

        <label>
          <strong>
            Documento PDF
          </strong>
        </label>

        <br />

        <input
          ref={fileInputRef}
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

            if (
              selectedFile.type !==
              'application/pdf'
            ) {

              setArchivo(null);

              setMessage(
                'Solo se permiten archivos PDF.'
              );

              event.target.value = '';

              return;
            }

            setMessage('');
            setArchivo(selectedFile);
          }}
          style={{
            marginTop: '6px'
          }}
        />

        {archivo && (

          <div
            style={{
              marginTop: '8px'
            }}
          >

            Archivo seleccionado:{' '}

            <strong>
              {archivo.name}
            </strong>

          </div>

        )}

      </div>



      <button
        type="button"
        disabled={saving}
        onClick={() => {
          void createRequerimiento();
        }}
        style={{
          padding: '10px 20px',
          cursor:
            saving
              ? 'not-allowed'
              : 'pointer'
        }}
      >

        {
          saving
            ? 'Enviando...'
            : 'Enviar requerimiento'
        }

      </button>



      {message && (

        <div
          style={{
            marginTop: '20px'
          }}
        >
          {message}
        </div>

      )}

    </div>

  );
};

export default Requerimientos;