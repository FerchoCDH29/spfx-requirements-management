import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import SharePointAdminService from './admin/SharePointAdminService';
import {
  ILookupOption,
  ISharePointFieldInfo,
  ISharePointItem,
  ISharePointListInfo,
  ISharePointUser,
  RestMethod
} from './admin/SharePointAdminModels';

export interface IAdminSharePointProps {
  context: WebPartContext;
}

type EditorMode = 'create' | 'edit' | null;

interface IOptionCache {
  [internalName: string]: ILookupOption[];
}

const AdminSharePoint: React.FC<IAdminSharePointProps> = ({ context }) => {
  const service = React.useMemo(
    () => new SharePointAdminService(context),
    [context]
  );

  const [lists, setLists] = React.useState<ISharePointListInfo[]>([]);
  const [selectedList, setSelectedList] =
    React.useState<ISharePointListInfo | undefined>(undefined);
  const [fields, setFields] = React.useState<ISharePointFieldInfo[]>([]);
  const [items, setItems] = React.useState<ISharePointItem[]>([]);
  const [users, setUsers] = React.useState<ISharePointUser[]>([]);
  const [lookupOptions, setLookupOptions] =
    React.useState<IOptionCache>({});
  const [loading, setLoading] = React.useState<boolean>(false);
  const [message, setMessage] = React.useState<string>('');
  const [search, setSearch] = React.useState<string>('');
  const [listSearch, setListSearch] = React.useState<string>('');

  const [editorMode, setEditorMode] =
    React.useState<EditorMode>(null);
  const [editingItem, setEditingItem] =
    React.useState<ISharePointItem | undefined>(undefined);
  const [formValues, setFormValues] =
    React.useState<Record<string, unknown>>({});

  const [selectedIds, setSelectedIds] =
    React.useState<number[]>([]);
  const [bulkField, setBulkField] = React.useState<string>('');
  const [bulkValue, setBulkValue] = React.useState<unknown>('');

  const [restMethod, setRestMethod] =
    React.useState<RestMethod>('GET');
  const [restEndpoint, setRestEndpoint] =
    React.useState<string>('/_api/web?$select=Title,Url');
  const [restBody, setRestBody] = React.useState<string>('');
  const [restResult, setRestResult] = React.useState<string>('');

  const editableFields = React.useMemo(
    () => fields.filter(field =>
      !field.Hidden &&
      !field.ReadOnlyField &&
      field.InternalName !== 'ID' &&
      field.InternalName !== 'Id' &&
      field.InternalName !== 'Created' &&
      field.InternalName !== 'Modified' &&
      field.InternalName !== 'Author' &&
      field.InternalName !== 'Editor' &&
      field.TypeAsString !== 'Computed'
    ),
    [fields]
  );

  const visibleFields = React.useMemo(
    () => fields.filter(field =>
      !field.Hidden &&
      field.InternalName !== 'Attachments' &&
      field.InternalName !== 'ContentType' &&
      field.TypeAsString !== 'Computed'
    ).slice(0, 12),
    [fields]
  );

  const filteredLists = React.useMemo(() => {
    const text = listSearch.trim().toLowerCase();
    if (!text) {
      return lists;
    }
    return lists.filter(list =>
      list.Title.toLowerCase().indexOf(text) >= 0
    );
  }, [lists, listSearch]);

  const filteredItems = React.useMemo(() => {
    const text = search.trim().toLowerCase();
    if (!text) {
      return items;
    }

    return items.filter(item =>
      Object.keys(item).some(key =>
        String(item[key] === undefined ? '' : item[key])
          .toLowerCase()
          .indexOf(text) >= 0
      )
    );
  }, [items, search]);

  const loadLists = React.useCallback(async (): Promise<void> => {
    try {
      setLoading(true);
      setMessage('');
      const result = await service.getLists();
      setLists(result);
      setMessage(`${result.length} listas disponibles.`);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : String(error));
    } finally {
      setLoading(false);
    }
  }, [service]);

  React.useEffect(() => {
    void loadLists();
  }, [loadLists]);

  const loadOptions = async (
    currentFields: ISharePointFieldInfo[]
  ): Promise<void> => {
    const cache: IOptionCache = {};
    let loadedUsers = users;

    const hasUser = currentFields.some(field =>
      field.TypeAsString === 'User' ||
      field.TypeAsString === 'UserMulti'
    );

    if (hasUser && loadedUsers.length === 0) {
      loadedUsers = await service.getUsers();
      setUsers(loadedUsers);
    }

    for (const field of currentFields) {
      if (
        field.TypeAsString === 'Lookup' &&
        field.LookupList &&
        field.LookupField
      ) {
        try {
          cache[field.InternalName] =
            await service.getLookupOptions(field);
        } catch {
          cache[field.InternalName] = [];
        }
      }
    }

    setLookupOptions(cache);
  };

  const selectList = async (
    list: ISharePointListInfo
  ): Promise<void> => {
    try {
      setLoading(true);
      setMessage('');
      setSelectedList(list);
      setSelectedIds([]);
      setEditorMode(null);
      setSearch('');

      const currentFields = await service.getFields(list.Title);
      const currentItems =
        await service.getItems(list.Title, currentFields);

      setFields(currentFields);
      setItems(currentItems);

      await loadOptions(currentFields);

      setMessage(
        `${list.Title}: ${currentItems.length} elementos cargados.`
      );
    } catch (error) {
      setMessage(error instanceof Error ? error.message : String(error));
    } finally {
      setLoading(false);
    }
  };

  const refreshSelectedList = async (): Promise<void> => {
    if (selectedList) {
      await selectList(selectedList);
    }
  };

  const getValueKey = (
    field: ISharePointFieldInfo
  ): string => {
    if (
      field.TypeAsString === 'User' ||
      field.TypeAsString === 'Lookup'
    ) {
      return `${field.InternalName}Id`;
    }

    return field.InternalName;
  };

  const beginCreate = (): void => {
    const values: Record<string, unknown> = {};

    editableFields.forEach(field => {
      if (field.TypeAsString === 'Boolean') {
        values[getValueKey(field)] = false;
      } else {
        values[getValueKey(field)] = '';
      }
    });

    setEditingItem(undefined);
    setFormValues(values);
    setEditorMode('create');
  };

  const beginEdit = (item: ISharePointItem): void => {
    const values: Record<string, unknown> = {};

    editableFields.forEach(field => {
      const key = getValueKey(field);
      values[key] = item[key] === undefined ? '' : item[key];
    });

    setEditingItem(item);
    setFormValues(values);
    setEditorMode('edit');
  };

  const normalizeBody = (
    values: Record<string, unknown>
  ): Record<string, unknown> => {
    const body: Record<string, unknown> = {};

    editableFields.forEach(field => {
      const key = getValueKey(field);
      let value = values[key];

      if (value === '' || value === undefined) {
        value = null;
      }

      if (
        value !== null &&
        (
          field.TypeAsString === 'Number' ||
          field.TypeAsString === 'Currency' ||
          field.TypeAsString === 'Integer' ||
          field.TypeAsString === 'Counter' ||
          field.TypeAsString === 'User' ||
          field.TypeAsString === 'Lookup'
        )
      ) {
        value = Number(value);
      }

      body[key] = value;
    });

    return body;
  };

  const saveEditor = async (): Promise<void> => {
    if (!selectedList || !editorMode) {
      return;
    }

    try {
      setLoading(true);
      setMessage('');

      const body = normalizeBody(formValues);

      if (editorMode === 'create') {
        await service.createItem(selectedList.Title, body);
        setMessage('Elemento creado correctamente.');
      } else if (editingItem) {
        await service.updateItem(
          selectedList.Title,
          editingItem.Id,
          body
        );
        setMessage(`Elemento #${editingItem.Id} actualizado.`);
      }

      setEditorMode(null);
      await refreshSelectedList();
    } catch (error) {
      setMessage(error instanceof Error ? error.message : String(error));
    } finally {
      setLoading(false);
    }
  };

  const deleteItem = async (item: ISharePointItem): Promise<void> => {
    if (!selectedList) {
      return;
    }

    const confirmed = window.confirm(
      `¿Eliminar definitivamente el elemento #${item.Id} de ${selectedList.Title}?`
    );

    if (!confirmed) {
      return;
    }

    try {
      setLoading(true);
      await service.deleteItem(selectedList.Title, item.Id);
      setMessage(`Elemento #${item.Id} eliminado.`);
      await refreshSelectedList();
    } catch (error) {
      setMessage(error instanceof Error ? error.message : String(error));
    } finally {
      setLoading(false);
    }
  };

  const toggleSelected = (id: number): void => {
    setSelectedIds(previous =>
      previous.indexOf(id) >= 0
        ? previous.filter(currentId => currentId !== id)
        : [...previous, id]
    );
  };

  const applyBulkEdit = async (): Promise<void> => {
    if (!selectedList || !bulkField || selectedIds.length === 0) {
      return;
    }

    const field = editableFields.find(current =>
      current.InternalName === bulkField
    );

    if (!field) {
      return;
    }

    const key = getValueKey(field);
    let value: unknown = bulkValue;

    if (
      field.TypeAsString === 'Number' ||
      field.TypeAsString === 'Currency' ||
      field.TypeAsString === 'Integer' ||
      field.TypeAsString === 'User' ||
      field.TypeAsString === 'Lookup'
    ) {
      value = bulkValue === '' ? null : Number(bulkValue);
    }

    const confirmed = window.confirm(
      `Se actualizarán ${selectedIds.length} elementos.\n\n` +
      `${field.Title} = ${String(value)}\n\n¿Continuar?`
    );

    if (!confirmed) {
      return;
    }

    try {
      setLoading(true);

      for (const id of selectedIds) {
        await service.updateItem(
          selectedList.Title,
          id,
          { [key]: value }
        );
      }

      setMessage(`${selectedIds.length} elementos actualizados.`);
      setSelectedIds([]);
      await refreshSelectedList();
    } catch (error) {
      setMessage(error instanceof Error ? error.message : String(error));
    } finally {
      setLoading(false);
    }
  };

  const executeRest = async (): Promise<void> => {
    if (
      restMethod !== 'GET' &&
      !window.confirm(
        `Vas a ejecutar ${restMethod} directamente contra SharePoint.\n¿Continuar?`
      )
    ) {
      return;
    }

    try {
      setLoading(true);
      const result = await service.executeRest(
        restEndpoint,
        restMethod,
        restBody
      );

      setRestResult(
        `HTTP ${result.status}\n\n` +
        JSON.stringify(result.data, null, 2)
      );
    } catch (error) {
      setRestResult(
        error instanceof Error ? error.message : String(error)
      );
    } finally {
      setLoading(false);
    }
  };

  const renderFieldEditor = (
    field: ISharePointFieldInfo,
    value: unknown,
    onChange: (newValue: unknown) => void
  ): React.ReactNode => {
    if (field.TypeAsString === 'Boolean') {
      return (
        <input
          type="checkbox"
          checked={Boolean(value)}
          onChange={event => onChange(event.target.checked)}
        />
      );
    }

    if (field.TypeAsString === 'Choice' && field.Choices) {
      return (
        <select
          value={String(value === null ? '' : value)}
          onChange={event => onChange(event.target.value)}
          style={{ width: '100%', padding: 8 }}
        >
          <option value="">-- seleccionar --</option>
          {field.Choices.map(choice => (
            <option key={choice} value={choice}>{choice}</option>
          ))}
        </select>
      );
    }

    if (field.TypeAsString === 'User') {
      return (
        <select
          value={String(value === null ? '' : value)}
          onChange={event => onChange(event.target.value)}
          style={{ width: '100%', padding: 8 }}
        >
          <option value="">-- sin usuario --</option>
          {users.map(user => (
            <option key={user.Id} value={user.Id}>
              {user.Title} {user.Email ? `(${user.Email})` : ''}
            </option>
          ))}
        </select>
      );
    }

    if (field.TypeAsString === 'Lookup') {
      const options = lookupOptions[field.InternalName] || [];

      return (
        <select
          value={String(value === null ? '' : value)}
          onChange={event => onChange(event.target.value)}
          style={{ width: '100%', padding: 8 }}
        >
          <option value="">-- sin valor --</option>
          {options.map(option => (
            <option key={option.Id} value={option.Id}>
              {option.Label}
            </option>
          ))}
        </select>
      );
    }

    if (field.TypeAsString === 'Note') {
      return (
        <textarea
          value={String(value === null ? '' : value)}
          onChange={event => onChange(event.target.value)}
          rows={4}
          style={{ width: '100%', padding: 8 }}
        />
      );
    }

    const type =
      field.TypeAsString === 'DateTime'
        ? 'datetime-local'
        : (
          field.TypeAsString === 'Number' ||
          field.TypeAsString === 'Currency' ||
          field.TypeAsString === 'Integer'
        )
          ? 'number'
          : 'text';

    return (
      <input
        type={type}
        value={String(value === null ? '' : value)}
        onChange={event => onChange(event.target.value)}
        style={{ width: '100%', padding: 8, boxSizing: 'border-box' }}
      />
    );
  };

  const formatCell = (
    item: ISharePointItem,
    field: ISharePointFieldInfo
  ): string => {
    const key = getValueKey(field);
    const value = item[key];

    if (value === null || value === undefined || value === '') {
      return '-';
    }

    if (field.TypeAsString === 'Boolean') {
      return value ? 'Sí' : 'No';
    }

    if (field.TypeAsString === 'User') {
      const user = users.find(current => current.Id === Number(value));
      return user ? user.Title : `Usuario #${String(value)}`;
    }

    if (field.TypeAsString === 'Lookup') {
      const options = lookupOptions[field.InternalName] || [];
      const option = options.find(current => current.Id === Number(value));
      return option ? option.Label : `ID ${String(value)}`;
    }

    if (Array.isArray(value)) {
      return value.join(', ');
    }

    return String(value);
  };

  const bulkFieldInfo = editableFields.find(field =>
    field.InternalName === bulkField
  );

  return (
    <div style={{
      fontFamily: 'Segoe UI, Arial, sans-serif',
      minHeight: 700,
      background: '#f6f7f9',
      color: '#242424'
    }}>
      <div style={{
        padding: '18px 24px',
        background: '#ffffff',
        borderBottom: '1px solid #e1e1e1',
        display: 'flex',
        justifyContent: 'space-between',
        alignItems: 'center'
      }}>
        <div>
          <h1 style={{ margin: 0, fontSize: 24 }}>SharePoint Admin</h1>
          <div style={{ marginTop: 4, color: '#666' }}>
            {service.getSiteUrl()}
          </div>
        </div>

        <button type="button" disabled={loading} onClick={() => void loadLists()}>
          Actualizar listas
        </button>
      </div>

      {message && (
        <div style={{
          margin: '14px 20px 0',
          padding: 10,
          background: '#fff',
          border: '1px solid #ddd'
        }}>
          {message}
        </div>
      )}

      <div style={{
        display: 'grid',
        gridTemplateColumns: '260px minmax(0, 1fr)',
        minHeight: 620
      }}>
        <aside style={{
          padding: 16,
          background: '#fff',
          borderRight: '1px solid #ddd'
        }}>
          <input
            placeholder="Buscar lista..."
            value={listSearch}
            onChange={event => setListSearch(event.target.value)}
            style={{
              width: '100%',
              padding: 9,
              boxSizing: 'border-box',
              marginBottom: 12
            }}
          />

          <div style={{ fontWeight: 600, marginBottom: 8 }}>
            Listas ({filteredLists.length})
          </div>

          {filteredLists.map(list => (
            <button
              key={list.Id}
              type="button"
              onClick={() => void selectList(list)}
              style={{
                display: 'block',
                width: '100%',
                textAlign: 'left',
                padding: '9px 10px',
                marginBottom: 4,
                border: 'none',
                cursor: 'pointer',
                background:
                  selectedList && selectedList.Id === list.Id
                    ? '#e8f2ff'
                    : 'transparent'
              }}
            >
              <strong>{list.Title}</strong>
              <div style={{ fontSize: 12, color: '#777' }}>
                {list.ItemCount} elementos
              </div>
            </button>
          ))}
        </aside>

        <main style={{ padding: 20, overflowX: 'auto' }}>
          {!selectedList && (
            <div style={{
              background: '#fff',
              padding: 30,
              border: '1px solid #ddd'
            }}>
              Selecciona una lista para consultar su estructura y sus datos.
            </div>
          )}

          {selectedList && (
            <>
              <div style={{
                display: 'flex',
                justifyContent: 'space-between',
                gap: 12,
                alignItems: 'center',
                marginBottom: 16
              }}>
                <div>
                  <h2 style={{ margin: 0 }}>{selectedList.Title}</h2>
                  <div style={{ color: '#666', marginTop: 4 }}>
                    {items.length} cargados · {fields.length} columnas detectadas
                  </div>
                </div>

                <div style={{ display: 'flex', gap: 8 }}>
                  <button
                    type="button"
                    disabled={loading}
                    onClick={() => void refreshSelectedList()}
                  >
                    Recargar
                  </button>
                  <button
                    type="button"
                    disabled={loading}
                    onClick={beginCreate}
                  >
                    + Nuevo elemento
                  </button>
                </div>
              </div>

              <div style={{
                background: '#fff',
                padding: 12,
                border: '1px solid #ddd',
                marginBottom: 14
              }}>
                <input
                  placeholder="Buscar en los elementos cargados..."
                  value={search}
                  onChange={event => setSearch(event.target.value)}
                  style={{
                    width: '100%',
                    padding: 9,
                    boxSizing: 'border-box'
                  }}
                />
              </div>

              {selectedIds.length > 0 && (
                <div style={{
                  background: '#fff8e5',
                  border: '1px solid #e6c86e',
                  padding: 12,
                  marginBottom: 14
                }}>
                  <strong>{selectedIds.length} seleccionados</strong>
                  <div style={{
                    display: 'grid',
                    gridTemplateColumns: '240px minmax(220px, 1fr) auto',
                    gap: 8,
                    marginTop: 10
                  }}>
                    <select
                      value={bulkField}
                      onChange={event => {
                        setBulkField(event.target.value);
                        setBulkValue('');
                      }}
                    >
                      <option value="">Campo a modificar...</option>
                      {editableFields.map(field => (
                        <option
                          key={field.InternalName}
                          value={field.InternalName}
                        >
                          {field.Title}
                        </option>
                      ))}
                    </select>

                    {bulkFieldInfo
                      ? renderFieldEditor(
                          bulkFieldInfo,
                          bulkValue,
                          setBulkValue
                        )
                      : <div />}

                    <button
                      type="button"
                      disabled={!bulkField || loading}
                      onClick={() => void applyBulkEdit()}
                    >
                      Aplicar
                    </button>
                  </div>
                </div>
              )}

              <div style={{
                background: '#fff',
                border: '1px solid #ddd',
                overflowX: 'auto'
              }}>
                <table style={{
                  width: '100%',
                  borderCollapse: 'collapse',
                  minWidth: 900
                }}>
                  <thead>
                    <tr style={{ background: '#f3f3f3' }}>
                      <th style={{ padding: 8 }} />
                      <th style={{ padding: 8, textAlign: 'left' }}>ID</th>
                      {visibleFields.map(field => (
                        <th
                          key={field.InternalName}
                          style={{ padding: 8, textAlign: 'left' }}
                        >
                          {field.Title}
                        </th>
                      ))}
                      <th style={{ padding: 8 }}>Acciones</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredItems.map(item => (
                      <tr key={item.Id} style={{ borderTop: '1px solid #eee' }}>
                        <td style={{ padding: 8 }}>
                          <input
                            type="checkbox"
                            checked={selectedIds.indexOf(item.Id) >= 0}
                            onChange={() => toggleSelected(item.Id)}
                          />
                        </td>
                        <td style={{ padding: 8 }}>{item.Id}</td>
                        {visibleFields.map(field => (
                          <td
                            key={field.InternalName}
                            style={{ padding: 8, maxWidth: 260 }}
                          >
                            {formatCell(item, field)}
                          </td>
                        ))}
                        <td style={{
                          padding: 8,
                          whiteSpace: 'nowrap',
                          textAlign: 'center'
                        }}>
                          <button
                            type="button"
                            onClick={() => beginEdit(item)}
                          >
                            Editar
                          </button>{' '}
                          <button
                            type="button"
                            onClick={() => void deleteItem(item)}
                          >
                            Eliminar
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <details style={{
                marginTop: 18,
                background: '#fff',
                border: '1px solid #ddd',
                padding: 14
              }}>
                <summary style={{ cursor: 'pointer', fontWeight: 600 }}>
                  Columnas detectadas ({fields.length})
                </summary>

                <div style={{ overflowX: 'auto', marginTop: 12 }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                    <thead>
                      <tr>
                        <th align="left">Nombre</th>
                        <th align="left">InternalName</th>
                        <th align="left">Tipo</th>
                        <th align="left">Editable</th>
                      </tr>
                    </thead>
                    <tbody>
                      {fields.map(field => (
                        <tr key={field.InternalName}>
                          <td>{field.Title}</td>
                          <td>{field.InternalName}</td>
                          <td>{field.TypeAsString}</td>
                          <td>{field.ReadOnlyField ? 'No' : 'Sí'}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </details>
            </>
          )}

          <details style={{
            marginTop: 18,
            background: '#151515',
            color: '#eee',
            padding: 14
          }}>
            <summary style={{ cursor: 'pointer', fontWeight: 600 }}>
              REST Console
            </summary>

            <div style={{
              display: 'grid',
              gridTemplateColumns: '120px 1fr',
              gap: 8,
              marginTop: 14
            }}>
              <select
                value={restMethod}
                onChange={event =>
                  setRestMethod(event.target.value as RestMethod)
                }
              >
                <option value="GET">GET</option>
                <option value="POST">POST</option>
                <option value="MERGE">MERGE</option>
                <option value="DELETE">DELETE</option>
              </select>

              <input
                value={restEndpoint}
                onChange={event => setRestEndpoint(event.target.value)}
                placeholder="/_api/..."
              />
            </div>

            <textarea
              value={restBody}
              onChange={event => setRestBody(event.target.value)}
              placeholder='Body JSON. Ej: {"Title":"Prueba"}'
              rows={6}
              style={{
                width: '100%',
                boxSizing: 'border-box',
                marginTop: 8,
                fontFamily: 'Consolas, monospace'
              }}
            />

            <button
              type="button"
              disabled={loading || !restEndpoint.trim()}
              onClick={() => void executeRest()}
              style={{ marginTop: 8 }}
            >
              Ejecutar
            </button>

            {restResult && (
              <pre style={{
                whiteSpace: 'pre-wrap',
                overflowX: 'auto',
                background: '#050505',
                padding: 12
              }}>
                {restResult}
              </pre>
            )}
          </details>
        </main>
      </div>

      {editorMode && selectedList && (
        <div style={{
          position: 'fixed',
          inset: 0,
          background: 'rgba(0,0,0,.45)',
          display: 'flex',
          justifyContent: 'flex-end',
          zIndex: 9999
        }}>
          <div style={{
            width: 520,
            maxWidth: '92vw',
            height: '100%',
            overflowY: 'auto',
            background: '#fff',
            padding: 22,
            boxSizing: 'border-box'
          }}>
            <h2 style={{ marginTop: 0 }}>
              {editorMode === 'create'
                ? `Nuevo elemento · ${selectedList.Title}`
                : `Editar #${editingItem ? editingItem.Id : ''}`}
            </h2>

            {editableFields.map(field => {
              const key = getValueKey(field);
              return (
                <div key={field.InternalName} style={{ marginBottom: 14 }}>
                  <label style={{
                    display: 'block',
                    fontWeight: 600,
                    marginBottom: 5
                  }}>
                    {field.Title}
                    {field.Required ? ' *' : ''}
                    <span style={{
                      fontWeight: 400,
                      color: '#777',
                      marginLeft: 6,
                      fontSize: 12
                    }}>
                      {field.TypeAsString}
                    </span>
                  </label>

                  {renderFieldEditor(
                    field,
                    formValues[key],
                    newValue => setFormValues(previous => ({
                      ...previous,
                      [key]: newValue
                    }))
                  )}
                </div>
              );
            })}

            <div style={{
              display: 'flex',
              justifyContent: 'flex-end',
              gap: 8,
              marginTop: 20
            }}>
              <button
                type="button"
                disabled={loading}
                onClick={() => setEditorMode(null)}
              >
                Cancelar
              </button>

              <button
                type="button"
                disabled={loading}
                onClick={() => void saveEditor()}
              >
                Guardar
              </button>
            </div>
          </div>
        </div>
      )}

      {loading && (
        <div style={{
          position: 'fixed',
          right: 20,
          bottom: 20,
          background: '#222',
          color: '#fff',
          padding: '10px 14px',
          zIndex: 10000
        }}>
          Procesando...
        </div>
      )}
    </div>
  );
};

export default AdminSharePoint;
