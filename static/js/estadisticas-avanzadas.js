/**
 * ESTADÍSTICAS AVANZADAS - JavaScript
 * Funcionalidades adicionales para el dashboard ejecutivo
 */

// === CONFIGURACIÓN GLOBAL ===
const StatsConfig = {
    animationDuration: 750,
    refreshInterval: 300000, // 5 minutos
    chartTypes: {
        line: 'line',
        bar: 'bar',
        doughnut: 'doughnut',
        pie: 'pie',
        radar: 'radar',
        polarArea: 'polarArea'
    },
    colors: {
        primary: '#FF9800',
        secondary: '#2196F3',
        success: '#4CAF50',
        danger: '#F44336',
        warning: '#FFC107',
        info: '#00BCD4',
        purple: '#9C27B0',
        pink: '#E91E63',
        teal: '#009688',
        indigo: '#3F51B5'
    }
};

// === UTILIDADES ===
const Utils = {
    /**
     * Formatea número como moneda
     */
    formatCurrency: function(value) {
        return new Intl.NumberFormat('es-MX', {
            style: 'currency',
            currency: 'MXN'
        }).format(value);
    },

    /**
     * Formatea número con separadores de miles
     */
    formatNumber: function(value) {
        return new Intl.NumberFormat('es-MX').format(value);
    },

    /**
     * Formatea fecha
     */
    formatDate: function(date, format = 'short') {
        const d = new Date(date);
        if (format === 'short') {
            return d.toLocaleDateString('es-MX', {
                day: '2-digit',
                month: '2-digit',
                year: 'numeric'
            });
        } else if (format === 'long') {
            return d.toLocaleDateString('es-MX', {
                weekday: 'long',
                year: 'numeric',
                month: 'long',
                day: 'numeric'
            });
        }
        return d.toLocaleDateString('es-MX');
    },

    /**
     * Calcula porcentaje
     */
    calculatePercentage: function(value, total) {
        if (total === 0) return 0;
        return ((value / total) * 100).toFixed(1);
    },

    /**
     * Genera gradiente de colores
     */
    generateGradient: function(ctx, color1, color2) {
        const gradient = ctx.createLinearGradient(0, 0, 0, 400);
        gradient.addColorStop(0, color1);
        gradient.addColorStop(1, color2);
        return gradient;
    },

    /**
     * Descarga datos como JSON
     */
    downloadJSON: function(data, filename) {
        const dataStr = JSON.stringify(data, null, 2);
        const dataBlob = new Blob([dataStr], { type: 'application/json' });
        const url = URL.createObjectURL(dataBlob);
        const link = document.createElement('a');
        link.href = url;
        link.download = filename + '.json';
        link.click();
        URL.revokeObjectURL(url);
    },

    /**
     * Muestra loading overlay
     */
    showLoading: function() {
        document.getElementById('loadingOverlay').classList.add('active');
    },

    /**
     * Oculta loading overlay
     */
    hideLoading: function() {
        document.getElementById('loadingOverlay').classList.remove('active');
    },

    /**
     * Muestra toast notification
     */
    showToast: function(message, type = 'info') {
        // Implementación básica - puede mejorarse con una librería como toastr
        alert(message);
    }
};

// === GESTOR DE GRÁFICOS ===
const ChartManager = {
    charts: {},

    /**
     * Crea o actualiza un gráfico
     */
    createOrUpdate: function(canvasId, config) {
        const canvas = document.getElementById(canvasId);
        if (!canvas) {
            console.error(`Canvas ${canvasId} no encontrado`);
            return null;
        }

        // Si ya existe, destruir primero
        if (this.charts[canvasId]) {
            this.charts[canvasId].destroy();
        }

        // Crear nuevo gráfico
        const ctx = canvas.getContext('2d');
        this.charts[canvasId] = new Chart(ctx, config);
        return this.charts[canvasId];
    },

    /**
     * Descarga gráfico como imagen
     */
    download: function(chartId, filename) {
        const chart = this.charts[chartId];
        if (!chart) {
            console.error(`Gráfico ${chartId} no encontrado`);
            return;
        }

        const url = chart.toBase64Image();
        const link = document.createElement('a');
        link.download = (filename || chartId) + '.png';
        link.href = url;
        link.click();
    },

    /**
     * Cambia tipo de gráfico dinámicamente
     */
    changeType: function(chartId, newType) {
        const chart = this.charts[chartId];
        if (!chart) {
            console.error(`Gráfico ${chartId} no encontrado`);
            return;
        }

        chart.config.type = newType;
        chart.update();
    },

    /**
     * Actualiza datos de un gráfico
     */
    updateData: function(chartId, newData) {
        const chart = this.charts[chartId];
        if (!chart) {
            console.error(`Gráfico ${chartId} no encontrado`);
            return;
        }

        chart.data = newData;
        chart.update('active');
    },

    /**
     * Destruye todos los gráficos
     */
    destroyAll: function() {
        Object.keys(this.charts).forEach(chartId => {
            if (this.charts[chartId]) {
                this.charts[chartId].destroy();
            }
        });
        this.charts = {};
    }
};

// === GESTOR DE FILTROS ===
const FilterManager = {
    currentFilters: {},

    /**
     * Aplica filtros
     */
    apply: function(filters) {
        this.currentFilters = { ...this.currentFilters, ...filters };

        // Construir URL con parámetros
        const params = new URLSearchParams(this.currentFilters);
        const url = window.location.pathname + '?' + params.toString();

        // Navegar con loading
        Utils.showLoading();
        window.location.href = url;
    },

    /**
     * Limpia filtros
     */
    clear: function() {
        this.currentFilters = {};
        Utils.showLoading();
        window.location.href = window.location.pathname;
    },

    /**
     * Obtiene valor de filtro actual
     */
    get: function(key) {
        return this.currentFilters[key];
    }
};

// === EXPORTADOR DE DATOS ===
const DataExporter = {
    /**
     * Exporta tabla a CSV
     */
    tableToCSV: function(tableId, filename) {
        const table = document.getElementById(tableId);
        if (!table) {
            console.error(`Tabla ${tableId} no encontrada`);
            return;
        }

        let csv = [];
        const rows = table.querySelectorAll('tr');

        rows.forEach(row => {
            const cols = row.querySelectorAll('td, th');
            const csvRow = [];
            cols.forEach(col => {
                // Limpiar texto y escapar comillas
                let text = col.textContent.replace(/"/g, '""');
                csvRow.push('"' + text + '"');
            });
            csv.push(csvRow.join(','));
        });

        const csvContent = csv.join('\n');
        const csvFile = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        link.download = (filename || 'export') + '.csv';
        link.href = URL.createObjectURL(csvFile);
        link.click();
        URL.revokeObjectURL(link.href);
    },

    /**
     * Exporta datos a Excel (usando biblioteca)
     */
    toExcel: function(data, filename) {
        // Requiere una librería como SheetJS
        console.log('Exportar a Excel requiere SheetJS');
        Utils.showToast('Función no implementada - Use los reportes oficiales', 'info');
    },

    /**
     * Imprime vista actual
     */
    print: function() {
        window.print();
    }
};

// === COMPARADOR DE PERÍODOS ===
const PeriodComparator = {
    /**
     * Compara dos períodos
     */
    compare: function(period1Data, period2Data) {
        const comparison = {
            difference: period2Data.total - period1Data.total,
            percentageChange: Utils.calculatePercentage(
                period2Data.total - period1Data.total,
                period1Data.total
            ),
            trend: period2Data.total > period1Data.total ? 'up' : 'down'
        };
        return comparison;
    },

    /**
     * Muestra comparación en UI
     */
    displayComparison: function(elementId, comparison) {
        const element = document.getElementById(elementId);
        if (!element) return;

        const icon = comparison.trend === 'up' ? '↑' : '↓';
        const colorClass = comparison.trend === 'up' ? 'positive' : 'negative';

        element.innerHTML = `
            <span class="kpi-change ${colorClass}">
                ${icon} ${comparison.percentageChange}%
                (${Utils.formatCurrency(Math.abs(comparison.difference))})
            </span>
        `;
    }
};

// === REFRESH AUTOMÁTICO ===
const AutoRefresh = {
    intervalId: null,
    isEnabled: false,

    /**
     * Inicia refresh automático
     */
    start: function(intervalMs = StatsConfig.refreshInterval) {
        if (this.isEnabled) return;

        this.intervalId = setInterval(() => {
            console.log('Auto-refresh: recargando datos...');
            this.refresh();
        }, intervalMs);

        this.isEnabled = true;
        console.log(`Auto-refresh iniciado (cada ${intervalMs / 1000}s)`);
    },

    /**
     * Detiene refresh automático
     */
    stop: function() {
        if (this.intervalId) {
            clearInterval(this.intervalId);
            this.intervalId = null;
            this.isEnabled = false;
            console.log('Auto-refresh detenido');
        }
    },

    /**
     * Refresca datos
     */
    refresh: function() {
        // Recargar solo los datos sin recargar toda la página
        // Implementar llamada AJAX aquí
        const params = new URLSearchParams(window.location.search);

        fetch(window.location.pathname + '.json?' + params.toString())
            .then(response => response.json())
            .then(data => {
                console.log('Datos actualizados:', data);
                // Actualizar gráficos y KPIs
                this.updateDashboard(data);
            })
            .catch(error => {
                console.error('Error al refrescar datos:', error);
            });
    },

    /**
     * Actualiza dashboard con nuevos datos
     */
    updateDashboard: function(data) {
        // Actualizar KPIs
        // Actualizar gráficos
        // etc.
        Utils.showToast('Datos actualizados', 'success');
    }
};

// === BÚSQUEDA Y FILTRADO EN TIEMPO REAL ===
const LiveSearch = {
    /**
     * Inicializa búsqueda en tabla
     */
    initTableSearch: function(inputId, tableId) {
        const input = document.getElementById(inputId);
        const table = document.getElementById(tableId);

        if (!input || !table) return;

        input.addEventListener('keyup', function() {
            const filter = this.value.toUpperCase();
            const rows = table.getElementsByTagName('tr');

            for (let i = 1; i < rows.length; i++) {
                const row = rows[i];
                const cells = row.getElementsByTagName('td');
                let found = false;

                for (let j = 0; j < cells.length; j++) {
                    const cell = cells[j];
                    if (cell) {
                        const textValue = cell.textContent || cell.innerText;
                        if (textValue.toUpperCase().indexOf(filter) > -1) {
                            found = true;
                            break;
                        }
                    }
                }

                row.style.display = found ? '' : 'none';
            }
        });
    }
};

// === GESTOR DE TEMAS ===
const ThemeManager = {
    currentTheme: 'dark',

    /**
     * Cambia entre tema oscuro y claro
     */
    toggle: function() {
        this.currentTheme = this.currentTheme === 'dark' ? 'light' : 'dark';
        document.body.classList.toggle('light-theme');

        // Guardar preferencia
        localStorage.setItem('theme', this.currentTheme);

        // Actualizar gráficos con nuevos colores
        this.updateChartColors();
    },

    /**
     * Carga tema guardado
     */
    loadSaved: function() {
        const saved = localStorage.getItem('theme');
        if (saved && saved !== this.currentTheme) {
            this.toggle();
        }
    },

    /**
     * Actualiza colores de gráficos
     */
    updateChartColors: function() {
        // Actualizar Chart.defaults con nuevos colores
        Chart.defaults.color = this.currentTheme === 'dark' ? '#b0b0b0' : '#333';
        Chart.defaults.borderColor = this.currentTheme === 'dark' ? '#333' : '#ddd';

        // Recrear gráficos
        // ChartManager.destroyAll();
        // initializeAllCharts(); // Esta función debe estar definida en tu página
    }
};

// === SHORTCUTS DE TECLADO ===
const KeyboardShortcuts = {
    init: function() {
        document.addEventListener('keydown', (e) => {
            // Ctrl/Cmd + R: Refresh
            if ((e.ctrlKey || e.metaKey) && e.key === 'r') {
                e.preventDefault();
                AutoRefresh.refresh();
            }

            // Ctrl/Cmd + P: Print
            if ((e.ctrlKey || e.metaKey) && e.key === 'p') {
                e.preventDefault();
                DataExporter.print();
            }

            // Ctrl/Cmd + E: Export
            if ((e.ctrlKey || e.metaKey) && e.key === 'e') {
                e.preventDefault();
                DataExporter.tableToCSV('detailsTable', 'estadisticas');
            }
        });
    }
};

// === INICIALIZACIÓN ===
document.addEventListener('DOMContentLoaded', function() {
    console.log('📊 Dashboard Ejecutivo de Estadísticas cargado');

    // Cargar tema guardado
    ThemeManager.loadSaved();

    // Inicializar shortcuts
    KeyboardShortcuts.init();

    // Ocultar loading
    Utils.hideLoading();

    // Opcional: iniciar auto-refresh (comentado por defecto)
    // AutoRefresh.start();
});

// Exponer API global
window.StatsAPI = {
    Utils,
    ChartManager,
    FilterManager,
    DataExporter,
    PeriodComparator,
    AutoRefresh,
    LiveSearch,
    ThemeManager
};