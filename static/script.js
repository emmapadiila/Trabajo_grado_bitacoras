'use strict';

(() => {
  // ========== CONSTANTES Y CONFIGURACIÓN ==========
  const ABORT_KEYS = { 
    buscar: 'buscar', 
    todos: 'todos', 
    stats: 'stats', 
    conexion: 'conexion', 
    exportar: 'exportar', 
    mutacion: 'mutacion' 
  };

  // ========== ESTADO GLOBAL ==========
  const state = {
    datos: [],
    editMode: false,
    proyectoEditando: null,
    aborters: {},
    lastFocus: null,
    estadisticasData: null
  };

  // ========== UTILIDADES ==========
  const $ = (sel, root = document) => root.querySelector(sel);
  
  const h = (tag, props = {}, ...children) => {
    const el = document.createElement(tag);
    if (props.class) el.className = props.class;
    if (props.text !== undefined) el.textContent = props.text;
    if (props.dataset) Object.assign(el.dataset, props.dataset);
    if (props.attrs) Object.entries(props.attrs).forEach(([k, v]) => v !== undefined && el.setAttribute(k, v));
    children.flat().forEach(c => {
      if (c == null) return;
      if (typeof c === 'string') el.appendChild(document.createTextNode(c));
      else el.appendChild(c);
    });
    return el;
  };

  const debounce = (fn, ms = 350) => {
    let t;
    return (...args) => {
      clearTimeout(t);
      t = setTimeout(() => fn(...args), ms);
    };
  };

  // ========== MANEJO DE PETICIONES ==========
  function withAbortKey(key) {
    const ctrl = new AbortController();
    if (state.aborters[key]) state.aborters[key].abort();
    state.aborters[key] = ctrl;
    return ctrl;
  }

  async function fetchJSON(url, { body, method = 'GET', headers = {}, abortKey = null } = {}) {
    const ctrl = abortKey ? withAbortKey(abortKey) : new AbortController();
    const timeoutId = setTimeout(() => ctrl.abort(), 15000);
    const opts = { method, headers, signal: ctrl.signal };
    if (body) opts.body = body;
    let res, data;
    try {
      res = await fetch(url, opts);
    } finally {
      clearTimeout(timeoutId);
    }
    const ct = res.headers.get('content-type') || '';
    if (ct.includes('application/json')) data = await res.json().catch(() => ({}));
    else data = await res.text();
    if (!res.ok) throw new Error((data && data.error) || res.statusText || 'Error de red');
    return data;
  }

  async function downloadFile(url, { body, method = 'GET', headers = {}, filename = 'descarga', abortKey = null } = {}) {
    const ctrl = abortKey ? withAbortKey(abortKey) : new AbortController();
    const timeoutId = setTimeout(() => ctrl.abort(), 30000);
    const opts = { method, headers, signal: ctrl.signal };
    if (body) opts.body = body;
    let res;
    try {
      res = await fetch(url, opts);
    } finally {
      clearTimeout(timeoutId);
    }
    if (!res.ok) {
      let err = 'Error de descarga';
      try {
        const j = await res.json();
        err = j.error || err;
      } catch (_) { }
      throw new Error(err);
    }
    const blob = await res.blob();
    const dlUrl = URL.createObjectURL(blob);
    const a = h('a', { attrs: { href: dlUrl, download: filename } });
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(dlUrl);
  }

  // ========== REFERENCIAS DOM ==========
  const el = {
    searchInput: $('#search-input'),
    searchBtn: $('#search-btn'),
    showAllBtn: $('#show-all-btn'),
    exportPdfBtn: $('#export-pdf-btn'),
    exportExcelBtn: $('#export-excel-btn'),
    refreshBtn: $('#refresh-btn'),
    searchResults: $('#search-results'),
    addForm: $('#add-form'),
    formMessage: $('#form-message'),
    connectionStatus: $('#connection-status'),
    cancelEditBtn: $('#cancel-edit-btn'),
    // Stats
    totalProyectos: $('#total-proyectos'),
    propuestasAprobadas: $('#propuestas-aprobadas'),
    trabajosFinales: $('#trabajos-finales'),
    totalAnteproyectos: $('#total-anteproyectos'),
    // Modal
    modal: $('#detail-modal'),
    modalContent: $('#modal-content'),
    closeModal: $('.close'),
  };

  // ========== ACCESIBILIDAD ==========
  function initA11y() {
    el.searchResults?.setAttribute('role', 'region');
    el.searchResults?.setAttribute('aria-live', 'polite');
    el.connectionStatus?.setAttribute('role', 'status');
    el.connectionStatus?.setAttribute('aria-live', 'polite');
    el.formMessage?.setAttribute('role', 'status');
    el.formMessage?.setAttribute('aria-live', 'polite');
    if (el.modal) {
      el.modal.setAttribute('role', 'dialog');
      el.modal.setAttribute('aria-modal', 'true');
      const title = el.modal.querySelector('h2');
      if (title && !title.id) title.id = 'modal-title';
      el.modal.setAttribute('aria-labelledby', 'modal-title');
    }
  }

  function setBusy(target, val) {
    target?.setAttribute('aria-busy', String(!!val));
  }

  // ========== SISTEMA DE MENSAJES ==========
  function showMessage(mensaje, tipo = 'info') {
    if (!el.formMessage) return;
    el.formMessage.textContent = mensaje;
    el.formMessage.className = `message ${tipo}`;
    el.formMessage.style.display = 'block';
    clearTimeout(showMessage._t);
    showMessage._t = setTimeout(() => {
      el.formMessage.style.display = 'none';
    }, 5000);
  }

  // ========== RENDERIZADO DE RESULTADOS ==========
  function showLoadingResults() {
    if (!el.searchResults) return;
    setBusy(el.searchResults, true);
    el.searchResults.innerHTML = '';
    el.searchResults.appendChild(
      h('div', { class: 'loading' },
        h('i', { class: 'fas fa-spinner fa-spin', attrs: { 'aria-hidden': 'true' } }),
        ' Buscando proyectos...'
      )
    );
  }

  function showNoResults() {
    if (!el.searchResults) return;
    setBusy(el.searchResults, false);
    el.searchResults.innerHTML = '';
    el.searchResults.appendChild(
      h('div', { class: 'no-results' },
        h('i', { class: 'fas fa-search', attrs: { 'aria-hidden': 'true' } }),
        h('p', { text: 'No se encontraron proyectos que coincidan con la búsqueda.' })
      )
    );
  }

  function renderResultados(resultados) {
    if (!el.searchResults) return;
    setBusy(el.searchResults, false);
    el.searchResults.innerHTML = '';
    if (!Array.isArray(resultados) || resultados.length === 0) return showNoResults();

    const frag = document.createDocumentFragment();
    resultados.forEach((proyecto, index) => {
      const proyectoNombre = proyecto['Proyecto/Articulo'] || 'Sin título';
      const programa = proyecto['Programa'] || 'No especificado';
      const estudiante1 = proyecto['Estudiante 1'] || 'No especificado';
      const estudiante2 = proyecto['Estudiante 2'] || '';
      const asesor = proyecto['Asesor'] || 'No especificado';
      const evaluadores = [proyecto['Evaluador 1'], proyecto['Evaluador 2'], proyecto['Evaluador 3']]
        .filter(Boolean).map(s => String(s).trim()).filter(Boolean);

      const propuesta = proyecto['Propuesta'] || 'No especificado';
      const anteproyecto = proyecto['Anteproyecto'] || 'No especificado';
      const trabajoFinal = proyecto['Trabajo final'] || 'No especificado';
      const fechaSustentacion = proyecto['Fecha sustentación'] || 'No especificado';
      const hora = proyecto['Hora'] || 'No especificado';
      const convocatoria = proyecto['Convocatoria'] || 'No especificado';
      const articuloMonografia = proyecto['ARTICULO/MONOGRAFIA'] || 'No especificado';
      const ano = proyecto['Año'] || 'No especificado';

      const details = h('div', { class: 'result-details' },
        row('Programa:', programa),
        row('Estudiante 1:', estudiante1),
        estudiante2 ? row('Estudiante 2:', estudiante2) : null,
        row('Asesor:', asesor),
        evaluadores.length ? row('Evaluadores:', evaluadores.join(', ')) : null,
        row('Hora:', hora),
        row('Fecha Sustentación:', fechaSustentacion),
        row('Convocatoria:', convocatoria),
        row('ARTICULO/MONOGRAFIA:', articuloMonografia),
        row('Año:', String(ano)),
        row('Propuesta:', estadoBadge(propuesta)),
        row('Anteproyecto:', estadoBadge(anteproyecto)),
        row('Trabajo Final:', estadoBadge(trabajoFinal))
      );

      const actions = h('div', { class: 'result-actions' },
        h('button', { class: 'btn-view', attrs: { type: 'button' }, dataset: { index } },
          h('i', { class: 'fas fa-eye', attrs: { 'aria-hidden': 'true' } }),
          ' Ver Detalles Completos'
        ),
        h('button', { class: 'btn-edit', attrs: { type: 'button' }, dataset: { index } },
          h('i', { class: 'fas fa-edit', attrs: { 'aria-hidden': 'true' } }),
          ' Editar Proyecto'
        ),
      );

      const card = h('article', {
        class: 'result-item',
        dataset: { index },
        attrs: { tabindex: '0', 'aria-label': `Proyecto ${proyectoNombre} del programa ${programa}` }
      }, h('h3', { text: proyectoNombre }), details, actions);

      frag.appendChild(card);
    });
    el.searchResults.appendChild(frag);
  }

  function row(label, value) {
    const wrapper = h('div', { class: 'detail-item' });
    wrapper.appendChild(h('span', { class: 'detail-label', text: label }));
    if (value instanceof HTMLElement) wrapper.appendChild(value);
    else wrapper.appendChild(h('span', { text: String(value) }));
    return wrapper;
  }

  function estadoBadge(estado) {
    return h('span', { class: `estado ${getEstadoClass(estado)}`, text: estado || 'No especificado' });
  }

  // ========== CONEXION Y DATOS ==========
  async function verificarConexion() {
    if (!el.connectionStatus) return;
    el.connectionStatus.className = 'connection-status';
    el.connectionStatus.innerHTML = '<i class="fas fa-sync-alt fa-spin" aria-hidden="true"></i> Verificando conexión...';
    try {
      const data = await fetchJSON('/verificar-conexion', { abortKey: ABORT_KEYS.conexion });
      if (data.estado === 'conectado') {
        el.connectionStatus.innerHTML = `<i class="fas fa-check-circle" aria-hidden="true"></i> ${data.mensaje}`;
        el.connectionStatus.className = 'connection-status connected';
      } else if (data.estado === 'parcial') {
        el.connectionStatus.innerHTML = `<i class="fas fa-exclamation-triangle" aria-hidden="true"></i> ${data.mensaje}`;
        el.connectionStatus.className = 'connection-status warning';
      } else {
        el.connectionStatus.innerHTML = `<i class="fas fa-times-circle" aria-hidden="true"></i> ${data.mensaje || 'Error de conexión'}`;
        el.connectionStatus.className = 'connection-status error';
      }
    } catch (e) {
      console.error(e);
      el.connectionStatus.innerHTML = '<i class="fas fa-times-circle" aria-hidden="true"></i> Error de conexión con el servidor';
      el.connectionStatus.className = 'connection-status error';
    }
  }

  async function cargarTodosLosProyectos({ silent = false } = {}) {
    if (!silent) showLoadingResults();
    try {
      const data = await fetchJSON('/mostrar_todos', { abortKey: ABORT_KEYS.todos });
      state.datos = Array.isArray(data.resultados) ? data.resultados : [];
      renderResultados(state.datos);
      if (!silent) showMessage(`Se muestran todos los proyectos (${state.datos.length} registros)`, 'info');
      renderEstadisticasFallback();
    } catch (e) {
      showNoResults();
      showMessage(String(e.message || 'Error de conexión con el servidor'), 'error');
      console.error('Error cargando proyectos:', e);
    }
  }

  async function buscarProyectos(termino) {
    const query = (termino ?? el.searchInput?.value ?? '').trim();
    if (!query) return showMessage('Escribe algo para buscar', 'warning');
    showLoadingResults();
    try {
      const data = await fetchJSON('/buscar', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ termino: query }),
        abortKey: ABORT_KEYS.buscar
      });
      state.datos = Array.isArray(data.resultados) ? data.resultados : [];
      renderResultados(state.datos);
      showMessage(`Se encontraron ${state.datos.length} resultados para "${query}"`, 'info');
      renderEstadisticasFallback();
    } catch (e) {
      showNoResults();
      showMessage(String(e.message || 'Error de conexion con el servidor'), 'error');
      console.error('Error buscando proyectos:', e);
    }
  }

  // ========== EXPORTACION ==========
  async function exportarPdf() {
    if (!state.datos || state.datos.length === 0) return showMessage('No hay datos para exportar', 'warning');
    showMessage('Generando PDF...', 'info');
    try {
      await downloadFile('/exportar_pdf', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ datos: state.datos }),
        filename: `proyectos_filtrados_${new Date().toISOString().split('T')[0]}.pdf`,
        abortKey: ABORT_KEYS.exportar
      });
      showMessage('PDF exportado correctamente', 'success');
    } catch (e) {
      showMessage(String(e.message || 'Error al exportar PDF'), 'error');
    }
  }

  async function exportarExcel() {
    showMessage('Generando Excel...', 'info');
    try {
      await downloadFile('/exportar_excel', {
        filename: `base_datos_completa_${new Date().toISOString().split('T')[0]}.xlsx`,
        abortKey: ABORT_KEYS.exportar
      });
      showMessage('Excel exportado correctamente', 'success');
    } catch (e) {
      showMessage(String(e.message || 'Error al exportar Excel'), 'error');
    }
  }

  // ========== FUNCIONALIDAD PARA CUADROS INTERACTIVOS ==========
  function initChartBoxes() {
    const chartBoxes = document.querySelectorAll('.chart-box');
    
    chartBoxes.forEach(box => {
      box.addEventListener('click', function(e) {
        // Si se hace clic en el header o en elementos dentro del header
        if (e.target.closest('.chart-header')) {
          const isActive = this.classList.contains('active');
          
          // Cerrar todos los demás cuadros
          chartBoxes.forEach(otherBox => {
            if (otherBox !== this) {
              otherBox.classList.remove('active');
            }
          });
          
          // Alternar el estado actual
          this.classList.toggle('active');
          
          // Si se activa y no tiene datos cargados, cargar los datos
          if (!isActive && this.classList.contains('active')) {
            const chartType = this.dataset.chart;
            loadChartData(chartType);
          }
        }
      });
    });
  }

  function loadChartData(chartType) {
    const chartContainer = document.getElementById(`chart-${chartType}`);
    
    if (!chartContainer) return;
    
    // Mostrar loading
    chartContainer.innerHTML = `
      <div class="loading">
        <i class="fas fa-spinner fa-spin"></i> Cargando datos...
      </div>
    `;
    
    // Si ya tenemos datos de estadísticas, renderizar directamente
    if (state.estadisticasData) {
      renderSpecificChart(chartType, state.estadisticasData);
    } else {
      // Si no, cargar estadísticas completas
      cargarEstadisticasCompletas().then(() => {
        renderSpecificChart(chartType, state.estadisticasData);
      }).catch(error => {
        chartContainer.innerHTML = `
          <div class="chart-error">
            <i class="fas fa-exclamation-triangle"></i>
            <p>Error cargando datos: ${error.message}</p>
          </div>
        `;
      });
    }
  }

  function renderSpecificChart(chartType, data) {
    switch (chartType) {
      case 'programas':
        renderChartProgramas(data.por_programa || {});
        break;
      case 'asesores':
        renderChartAsesores(data.por_asesor || {});
        break;
      case 'propuestas':
        renderChartPropuestas(data.estados?.propuestas || {});
        break;
      case 'fechas':
        renderChartFechas(data.por_fecha || {});
        break;
    }
  }

  // ========== DASHBOARD DE ESTADÍSTICAS ==========
  async function cargarEstadisticasCompletas() {
    try {
      showMessage('Cargando estadisticas...', 'info');

      const data = await fetchJSON('/estadisticas-detalladas', {
        abortKey: ABORT_KEYS.stats
      });

      // Verificar si hay error en la respuesta
      if (data.error) {
        throw new Error(data.error);
      }

      // Guardar datos en el estado
      state.estadisticasData = data;

      // Actualizar tarjetas de resumen CON LOS NUEVOS CONTADORES
      if (el.totalProyectos) {
        el.totalProyectos.textContent = data.totales?.total_proyectos || 0;
      }

      if (el.propuestasAprobadas) {
        // Usar el nuevo contador específico en lugar del de graficos
        el.propuestasAprobadas.textContent = data.estados_contadores?.propuestas_aprobadas || 0;
      }

      if (el.trabajosFinales) {
        // Usar el nuevo contador específico en lugar del de gráficos
        el.trabajosFinales.textContent = data.estados_contadores?.trabajos_finales_aprobados || 0;
      }

      // Actualizar total anteproyectos
      if (el.totalAnteproyectos) {
        el.totalAnteproyectos.textContent = data.totales?.total_anteproyectos || 0;
      }

      // Si hay cuadros activos, actualizar sus graficas
      const activeChartBox = document.querySelector('.chart-box.active');
      if (activeChartBox) {
        const chartType = activeChartBox.dataset.chart;
        renderSpecificChart(chartType, data);
      }

      // Actualizar 
      const updateTime = document.getElementById('last-update');
      if (updateTime) {
        updateTime.textContent = new Date().toLocaleString();
      }

      showMessage('Estadísticas actualizadas correctamente', 'success');

    } catch (e) {
      console.error('Error cargando estadisticas completas:', e);
      showMessage('Error cargando estadisticas: ' + e.message, 'error');

      // Mostrar mensajes de error en los graficos activos
      const activeChartBox = document.querySelector('.chart-box.active');
      if (activeChartBox) {
        const chartType = activeChartBox.dataset.chart;
        const chartContainer = document.getElementById(`chart-${chartType}`);
        if (chartContainer) {
          chartContainer.innerHTML = `
            <div class="chart-error">
              <i class="fas fa-exclamation-triangle"></i>
              <p>Error cargando datos</p>
            </div>
          `;
        }
      }
    }
  }

  function renderChartProgramas(datos) {
    const container = document.getElementById('chart-programas');
    if (!container) return;

    const entries = Object.entries(datos);
    if (entries.length === 0) {
      container.innerHTML = '<div class="no-data">No hay datos de programas</div>';
      return;
    }

    const total = Object.values(datos).reduce((sum, val) => sum + val, 0);
    const colores = ['#B71C1C', '#1A4B8C', '#2E7D32', '#F57C00', '#6A1B9A', '#00838F', '#5D4037'];

    let html = '<div class="chart-container-simple">';

    entries.forEach(([programa, cantidad], index) => {
      const porcentaje = total > 0 ? ((cantidad / total) * 100).toFixed(1) : 0;
      const color = colores[index % colores.length];

      const programaCorto = programa.length > 25 ? programa.substring(0, 25) + '...' : programa;

      html += `
            <div class="chart-item">
                <div class="chart-bar-horizontal">
                    <div class="chart-bar-label">
                        <span class="chart-color" style="background: ${color}"></span>
                        ${programaCorto}
                    </div>
                    <div class="chart-bar-track">
                        <div class="chart-bar-fill" style="width: ${porcentaje}%; background: ${color}"></div>
                    </div>
                    <div class="chart-bar-value">${cantidad} (${porcentaje}%)</div>
                </div>
            </div>
        `;
    });

    html += '</div>';
    container.innerHTML = html;
  }

  function renderChartAsesores(datos) {
    const container = document.getElementById('chart-asesores');
    if (!container) return;

    const entries = Object.entries(datos);
    if (entries.length === 0) {
      container.innerHTML = '<div class="no-data">No hay datos de asesores</div>';
      return;
    }

    const max = Math.max(...Object.values(datos));
    let html = '<div class="chart-container-simple">';

    entries.forEach(([asesor, cantidad]) => {
      const porcentaje = max > 0 ? (cantidad / max) * 100 : 0;
      // Acortar nombres largos
      const asesorCorto = asesor.length > 30 ? asesor.substring(0, 30) + '...' : asesor;

      html += `
            <div class="chart-item">
                <div class="chart-bar-horizontal">
                    <div class="chart-bar-label">${asesorCorto}</div>
                    <div class="chart-bar-track">
                        <div class="chart-bar-fill" style="width: ${porcentaje}%; background: var(--unilibre-gold)"></div>
                    </div>
                    <div class="chart-bar-value">${cantidad}</div>
                </div>
            </div>
        `;
    });

    html += '</div>';
    container.innerHTML = html;
  }

  function renderChartPropuestas(datos) {
    const container = document.getElementById('chart-propuestas');
    if (!container) return;

    const estados = {
      'Aprobadas': datos.aprobados || 0,
      'En Revision': datos.revision || 0,
      'No Aprobadas': datos.no_aprobados || 0,
      'No Especificado': datos.no_especificado || 0
    };

    const total = Object.values(estados).reduce((sum, val) => sum + val, 0);
    if (total === 0) {
      container.innerHTML = '<div class="no-data">No hay datos de propuestas</div>';
      return;
    }

    const colores = {
      'Aprobadas': '#4CAF50',
      'En Revisión': '#FFC107',
      'No Aprobadas': '#F44336',
      'No Especificado': '#9E9E9E'
    };

    let html = '<div class="chart-container-simple">';

    Object.entries(estados).forEach(([estado, cantidad]) => {
      if (cantidad > 0) {
        const porcentaje = ((cantidad / total) * 100).toFixed(1);
        const color = colores[estado];

        html += `
                <div class="chart-item">
                    <div class="chart-bar-horizontal">
                        <div class="chart-bar-label">
                            <span class="chart-color" style="background: ${color}"></span>
                            ${estado}
                        </div>
                        <div class="chart-bar-track">
                            <div class="chart-bar-fill" style="width: ${porcentaje}%; background: ${color}"></div>
                        </div>
                        <div class="chart-bar-value">${cantidad} (${porcentaje}%)</div>
                    </div>
                </div>
            `;
      }
    });

    html += '</div>';
    container.innerHTML = html;
  }

  function renderChartFechas(datos) {
    const container = document.getElementById('chart-fechas');
    if (!container) return;

    const entries = Object.entries(datos);
    if (entries.length === 0) {
      container.innerHTML = '<div class="no-data">No hay datos por fecha</div>';
      return;
    }

    const max = Math.max(...Object.values(datos));
    let html = '<div class="chart-container-simple">';

    entries.forEach(([fecha, cantidad]) => {
      const porcentaje = max > 0 ? (cantidad / max) * 100 : 0;

      html += `
            <div class="chart-item">
                <div class="chart-bar-horizontal">
                    <div class="chart-bar-label">${fecha}</div>
                    <div class="chart-bar-track">
                        <div class="chart-bar-fill" style="width: ${porcentaje}%; background: var(--unilibre-red)"></div>
                    </div>
                    <div class="chart-bar-value">${cantidad}</div>
                </div>
            </div>
        `;
    });

    html += '</div>';
    container.innerHTML = html;
  }

  function renderEstadisticasFallback() {
    if (!Array.isArray(state.datos) || state.datos.length === 0) return;
    const toLower = v => String(v || '').toLowerCase();
    const propuestasAprobadas = state.datos.filter(p => toLower(p['Propuesta']).includes('aprobado')).length;
    const finalesAprobados = state.datos.filter(p => toLower(p['Trabajo final']).includes('aprobado')).length;
    if (el.propuestasAprobadas) el.propuestasAprobadas.textContent = propuestasAprobadas;
    if (el.trabajosFinales) el.trabajosFinales.textContent = finalesAprobados;
    if (el.totalProyectos) el.totalProyectos.textContent = state.datos.length;
  }

  // ========== EDICION DE PROYECTOS ==========
  function setValue(selector, value) {
    const input = $(selector);
    if (input) input.value = value || '';
  }

  function cargarDatosEnFormulario(proyecto) {
    state.editMode = true;
    state.proyectoEditando = proyecto;
    const submitText = $('#submit-text');
    const modoEdicionInput = $('#modo_edicion');
    const numeroFilaInput = $('#numero_fila');
    const cancelEditBtn = el.cancelEditBtn;

    if (submitText) submitText.textContent = 'Actualizar Proyecto';
    if (modoEdicionInput) modoEdicionInput.value = 'editar';
    if (numeroFilaInput) numeroFilaInput.value = proyecto.numero_fila || '';
    if (cancelEditBtn) cancelEditBtn.style.display = 'block';

    setValue('#proyecto', proyecto['Proyecto/Articulo']);
    setValue('#programa', proyecto['Programa']);
    setValue('#estudiante1', proyecto['Estudiante 1']);
    setValue('#estudiante2', proyecto['Estudiante 2']);
    setValue('#asesor', proyecto['Asesor']);
    setValue('#evaluador1', proyecto['Evaluador 1']);
    setValue('#evaluador2', proyecto['Evaluador 2']);
    setValue('#evaluador3', proyecto['Evaluador 3']);
    setValue('#hora', proyecto['Hora']);

    const fecha = convertirFechaFormatoInput(proyecto['Fecha sustentación'] || '');
    setValue('#fecha', fecha || proyecto['Fecha sustentación'] || '');

    setValue('#convocatoria', proyecto['Convocatoria']);
    setValue('#articulo_monografia', proyecto['ARTICULO/MONOGRAFIA']);
    setValue('#ano', proyecto['Año']);
    setValue('#propuesta', proyecto['Propuesta']);
    setValue('#anteproyecto', proyecto['Anteproyecto']);
    setValue('#trabajo-final', proyecto['Trabajo final']);

    document.querySelector('.add-section')?.scrollIntoView({ behavior: 'smooth' });
    showMessage('Modo edicion activado. Modifica los campos y haz clic en "Actualizar Proyecto"', 'info');
  }

  function cancelarEdicion() {
    state.editMode = false;
    state.proyectoEditando = null;
    const submitText = $('#submit-text');
    const modoEdicionInput = $('#modo_edicion');
    const cancelEditBtn = el.cancelEditBtn;
    if (submitText) submitText.textContent = 'Guardar en Google Sheets';
    if (modoEdicionInput) modoEdicionInput.value = 'agregar';
    if (cancelEditBtn) cancelEditBtn.style.display = 'none';
    el.addForm?.reset();
    showMessage('Edición cancelada', 'info');
  }

  async function agregarProyecto(e) {
    e.preventDefault();
    if (!el.addForm) return;
    const formData = new FormData(el.addForm);
    const datos = Object.fromEntries(formData.entries());
    const modo = ($('#modo_edicion')?.value) || 'agregar';
    if (!datos.proyecto_articulo || !datos.estudiante1) return showMessage('Los campos Proyecto/Artículo y Estudiante 1 son obligatorios', 'error');

    const submitBtn = $('#submit-btn');
    if (!submitBtn) return;
    const original = submitBtn.innerHTML;
    submitBtn.innerHTML = `<i class="fas fa-spinner fa-spin" aria-hidden="true"></i> ${modo === 'editar' ? 'Actualizando...' : 'Guardando...'}`;
    submitBtn.disabled = true;

    try {
      const url = modo === 'editar' ? '/actualizar' : '/agregar';
      if (modo === 'editar') {
        const numeroFilaInput = $('#numero_fila');
        if (numeroFilaInput) datos.numero_fila = numeroFilaInput.value;
      }
      await fetchJSON(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(datos),
        abortKey: ABORT_KEYS.mutacion
      });
      if (modo === 'editar') {
        showMessage('Proyecto actualizado exitosamente en Google Sheets', 'success');
        cancelarEdicion();
      } else {
        showMessage('Proyecto agregado exitosamente a Google Sheets', 'success');
        el.addForm.reset();
      }
      await Promise.all([cargarEstadisticasCompletas(), cargarTodosLosProyectos({ silent: true })]);
      renderResultados(state.datos);
    } catch (e) {
      showMessage(String(e.message || `Error al ${modo === 'editar' ? 'actualizar' : 'agregar'} el proyecto`), 'error');
      console.error(e);
    } finally {
      submitBtn.innerHTML = original;
      submitBtn.disabled = false;
    }
  }

  // ========== MODAL ==========
  function buildModalContent(proyecto) {
    const proyectoNombre = proyecto['Proyecto/Articulo'] || 'Sin título';
    const programa = proyecto['Programa'] || 'No especificado';
    const estudiante1 = proyecto['Estudiante 1'] || 'No especificado';
    const estudiante2 = proyecto['Estudiante 2'] || 'No especificado';
    const asesor = proyecto['Asesor'] || 'No especificado';
    const evaluador1 = proyecto['Evaluador 1'] || 'No especificado';
    const evaluador2 = proyecto['Evaluador 2'] || 'No especificado';
    const evaluador3 = proyecto['Evaluador 3'] || 'No especificado';
    const hora = proyecto['Hora'] || 'No especificado';
    const fechaSustentacion = proyecto['Fecha sustentación'] || 'No especificado';
    const propuesta = proyecto['Propuesta'] || 'No especificado';
    const anteproyecto = proyecto['Anteproyecto'] || 'No especificado';
    const trabajoFinal = proyecto['Trabajo final'] || 'No especificado';
    const convocatoria = proyecto['Convocatoria'] || 'No especificado';
    const articuloMonografia = proyecto['ARTICULO/MONOGRAFIA'] || 'No especificado';
    const ano = proyecto['Año'] || 'No especificado';
    const hojaOrigen = proyecto['hoja_origen'] || 'No especificado';

    const container = h('div', { class: 'project-details' });
    const title = el.modal.querySelector('h2');
    if (title) title.textContent = 'Detalles del Proyecto';
    const grid = h('div', { class: 'details-grid' },
      detailRow('Programa:', programa),
      detailRow('Estudiante 1:', estudiante1),
      detailRow('Estudiante 2:', estudiante2),
      detailRow('Asesor:', asesor),
      detailRow('Evaluador 1:', evaluador1),
      detailRow('Evaluador 2:', evaluador2),
      detailRow('Evaluador 3:', evaluador3),
      detailRow('Hora:', hora),
      detailRow('Fecha Sustentación:', fechaSustentacion),
      detailRow('Convocatoria:', convocatoria),
      detailRow('ARTICULO/MONOGRAFIA:', articuloMonografia),
      detailRow('Año:', String(ano)),
      detailRow('Propuesta:', estadoBadge(propuesta)),
      detailRow('Anteproyecto:', estadoBadge(anteproyecto)),
      detailRow('Trabajo Final:', estadoBadge(trabajoFinal)),
      detailRow('Hoja Origen:', hojaOrigen),
    );
    container.appendChild(h('h3', { text: proyectoNombre }));
    container.appendChild(grid);
    return container;
  }

  function detailRow(label, value) {
    const row = h('div', { class: 'detail-row' });
    row.appendChild(h('strong', { text: label }));
    if (value instanceof HTMLElement) row.appendChild(value);
    else row.appendChild(h('span', { text: String(value) }));
    return row;
  }

  function abrirModal(proyecto) {
    if (!el.modal || !el.modalContent) return;
    state.lastFocus = document.activeElement;
    el.modalContent.innerHTML = '';
    el.modalContent.appendChild(buildModalContent(proyecto));
    el.modal.style.display = 'block';
    el.modal.removeAttribute('aria-hidden');
    trapFocus(el.modal);
    el.closeModal?.focus();
  }

  function cerrarModal() {
    if (!el.modal) return;
    el.modal.style.display = 'none';
    el.modal.setAttribute('aria-hidden', 'true');
    releaseFocusTrap();
    state.lastFocus?.focus?.();
  }

  let activeTrap = null;

  function trapFocus(container) {
    const focusable = 'a, button, textarea, input, select, [tabindex]:not([tabindex="-1"])';
    const nodes = Array.from(container.querySelectorAll(focusable)).filter(n => !n.hasAttribute('disabled'));
    if (!nodes.length) return;
    const first = nodes[0], last = nodes[nodes.length - 1];
    function onKeydown(e) {
      if (e.key === 'Escape') {
        e.preventDefault();
        return cerrarModal();
      }
      if (e.key !== 'Tab') return;
      if (e.shiftKey && document.activeElement === first) {
        e.preventDefault();
        last.focus();
      } else if (!e.shiftKey && document.activeElement === last) {
        e.preventDefault();
        first.focus();
      }
    }
    activeTrap = onKeydown;
    container.addEventListener('keydown', onKeydown);
  }

  function releaseFocusTrap() {
    if (activeTrap && el.modal) {
      el.modal.removeEventListener('keydown', activeTrap);
      activeTrap = null;
    }
  }

  // ========== HELPERS ==========
  function convertirFechaFormatoInput(fechaString) {
    if (!fechaString) return '';
    const formatos = [
      /(\d{1,2})\s+de\s+(\w+)\s+de\s+(\d{4})/,
      /(\d{1,2})\/(\d{1,2})\/(\d{4})/,
      /(\d{4})-(\d{1,2})-(\d{1,2})/
    ];
    const meses = {
      'enero': '01', 'febrero': '02', 'marzo': '03', 'abril': '04', 'mayo': '05', 'junio': '06',
      'julio': '07', 'agosto': '08', 'septiembre': '09', 'octubre': '10', 'noviembre': '11', 'diciembre': '12'
    };
    for (const formato of formatos) {
      const m = fechaString.match(formato);
      if (m) {
        if (formato.source.includes('de')) {
          const d = m[1].padStart(2, '0');
          const mes = meses[String(m[2]).toLowerCase()] || '01';
          return `${m[3]}-${mes}-${d}`;
        }
        if (formato.source.includes('/')) {
          const d = m[1].padStart(2, '0');
          const mes = m[2].padStart(2, '0');
          return `${m[3]}-${mes}-${d}`;
        }
        return fechaString;
      }
    }
    return '';
  }

  function getEstadoClass(estado) {
    if (!estado) return 'estado-default';
    const s = String(estado).toLowerCase();
    if (s.includes('aprobado')) return 'estado-aprobado';
    if (s.includes('revisión') || s.includes('revision')) return 'estado-revision';
    if (s.includes('no aprobado') || s.includes('rechazado')) return 'estado-rechazado';
    if (s.includes('no aplica') || s.includes('n/a')) return 'estado-na';
    return 'estado-default';
  }

  // ========== EVENTOS ==========
  function bindEvents() {
    // Búsqueda
    el.searchBtn?.addEventListener('click', () => buscarProyectos());
    el.searchInput?.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') buscarProyectos();
    });
    el.searchInput?.addEventListener('input', debounce((e) => {
      const val = e.target.value.trim();
      if (val.length >= 3) buscarProyectos(val);
      else if (val.length === 0) cargarTodosLosProyectos();
    }, 350));

    // Botones principales
    el.showAllBtn?.addEventListener('click', () => cargarTodosLosProyectos());
    el.refreshBtn?.addEventListener('click', () => {
      verificarConexion();
      cargarEstadisticasCompletas();
      const q = el.searchInput?.value.trim();
      if (q) buscarProyectos(q);
      else cargarTodosLosProyectos();
    });

    // Exportación
    el.exportPdfBtn?.addEventListener('click', exportarPdf);
    el.exportExcelBtn?.addEventListener('click', exportarExcel);

    // Formulario
    el.addForm?.addEventListener('submit', agregarProyecto);
    el.cancelEditBtn?.addEventListener('click', cancelarEdicion);

    // Estadísticas
    const refreshStatsBtn = document.getElementById('refresh-stats-btn');
    if (refreshStatsBtn) {
      refreshStatsBtn.addEventListener('click', cargarEstadisticasCompletas);
    }

    // Delegación en resultados
    el.searchResults?.addEventListener('click', (e) => {
      const viewBtn = e.target.closest('.btn-view');
      const editBtn = e.target.closest('.btn-edit');
      const card = e.target.closest('.result-item');
      const targetEl = viewBtn || editBtn || card;
      if (!targetEl) return;
      const idx = Number(targetEl.dataset.index ?? card?.dataset.index);
      const proyecto = state.datos[idx];
      if (!proyecto) return;
      if (viewBtn || (!viewBtn && !editBtn && card)) abrirModal(proyecto);
      else if (editBtn) cargarDatosEnFormulario(proyecto);
    });

    el.searchResults?.addEventListener('keydown', (e) => {
      if ((e.key === 'Enter' || e.key === ' ') && e.target.classList.contains('.result-item')) {
        const idx = Number(e.target.dataset.index);
        const proyecto = state.datos[idx];
        if (proyecto) abrirModal(proyecto);
      }
    });

    // Modal
    el.closeModal?.addEventListener('click', cerrarModal);
    window.addEventListener('click', (e) => {
      if (e.target === el.modal) cerrarModal();
    });
    window.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && el.modal?.style.display === 'block') cerrarModal();
    });
  }

  // ========== INICIALIZACIÓN ==========
  document.addEventListener('DOMContentLoaded', () => {
    initA11y();
    bindEvents();
    initChartBoxes(); // Inicializar cuadros interactivos
    verificarConexion();
    cargarEstadisticasCompletas();
    cargarTodosLosProyectos();
  });
})();