(function (global) {
  'use strict';

  const DEFAULT_LOCALE = 'es-MX';
  const STORAGE_TOKEN_KEY = 'seguimiento_cargas_token';
  const STORAGE_USER_KEY = 'seguimiento_cargas_user';
  const STORAGE_THEME_KEY = 'seguimiento_cargas_theme';
  const THEME_LIGHT = 'light';
  const THEME_DARK = 'dark';
  const DATE_HEADER_REGEX = /(fecha|cita|llegada|salida|hora)/i;
  const COLUMN_CONFIG = [
    { key: 'trip', label: 'Trip' },
    { key: 'ejecutivo', label: 'Ejecutivo' },
    { key: 'caja', label: 'Caja' },
    { key: 'referencia', label: 'Referencia' },
    { key: 'cliente', label: 'Cliente' },
    { key: 'destino', label: 'Destino' },
    { key: 'estatus', label: 'Estatus' },
    { key: 'segmento', label: 'Segmento' },
    { key: 'trmx', label: 'TR-MX' },
    { key: 'trusa', label: 'TR-USA' },
    { key: 'citaCarga', label: 'Cita carga' },
    { key: 'llegadaCarga', label: 'Llegada carga' },
    { key: 'citaEntrega', label: 'Cita entrega' },
    { key: 'llegadaEntrega', label: 'Llegada entrega' },
    { key: 'comentarios', label: 'Comentarios' },
    { key: 'docs', label: 'Docs' },
    { key: 'tracking', label: 'Tracking' }
  ];
  const DATE_FIELD_KEYS = ['citaCarga', 'llegadaCarga', 'citaEntrega', 'llegadaEntrega'];
  const DATE_FIELD_SET = new Set(DATE_FIELD_KEYS);
  const NOWRAP_COLUMN_KEYS = new Set(['trip', 'caja', 'trmx', 'trusa']);
  const MS_PER_DAY = 24 * 60 * 60 * 1000;
  const COLUMN_LABEL_TO_KEY = COLUMN_CONFIG.reduce(function (acc, column) {
    if (column && column.label) {
      acc[String(column.label).trim().toLowerCase()] = column.key;
    }
    return acc;
  }, {});
  const DEFAULT_USER = {
    id: 'admin',
    username: 'admin',
    password: 'admin123',
    displayName: 'Administrador'
  };

  function pad2(value) {
    const num = Number(value);
    if (!Number.isFinite(num)) {
      return '00';
    }
    const abs = Math.abs(Math.trunc(num));
    return abs < 10 ? `0${abs}` : String(abs);
  }

  function parseYear(value) {
    const year = parseInt(value, 10);
    if (!Number.isFinite(year)) {
      return NaN;
    }
    if (String(value).length === 2) {
      return year >= 70 ? 1900 + year : 2000 + year;
    }
    return year;
  }

  function parseDateParts(input) {
    if (input == null || input === '') {
      return null;
    }

    if (input instanceof Date && !isNaN(input.getTime())) {
      return {
        year: input.getFullYear(),
        month: input.getMonth() + 1,
        day: input.getDate(),
        hour: input.getHours(),
        minute: input.getMinutes()
      };
    }

    if (typeof input === 'number' && Number.isFinite(input)) {
      const date = new Date(input);
      if (!isNaN(date.getTime())) {
        return {
          year: date.getFullYear(),
          month: date.getMonth() + 1,
          day: date.getDate(),
          hour: date.getHours(),
          minute: date.getMinutes()
        };
      }
    }

    const value = String(input).trim();
    if (!value) {
      return null;
    }

    const isoMatch = value.match(/^([0-9]{4})-([0-9]{2})-([0-9]{2})[T\s]([0-9]{2}):([0-9]{2})(?::([0-9]{2}))?(?:\.[0-9]+)?(?:Z|[+-][0-9]{2}:?[0-9]{2})?$/i);
    if (isoMatch) {
      return {
        year: parseYear(isoMatch[1]),
        month: parseInt(isoMatch[2], 10),
        day: parseInt(isoMatch[3], 10),
        hour: parseInt(isoMatch[4], 10),
        minute: parseInt(isoMatch[5], 10)
      };
    }

    const dmyMatch = value.match(/^([0-9]{1,2})[\/\-.]([0-9]{1,2})[\/\-.]([0-9]{2,4})(?:[ T]([0-9]{1,2}):([0-9]{2})(?::([0-9]{2}))?(?:\s*([AP])M)?)?$/i);
    if (dmyMatch) {
      let hour = dmyMatch[4] != null ? parseInt(dmyMatch[4], 10) : 0;
      const minute = dmyMatch[5] != null ? parseInt(dmyMatch[5], 10) : 0;
      const meridiem = dmyMatch[7] ? dmyMatch[7].toUpperCase() : '';
      if (meridiem === 'P' && hour < 12) {
        hour += 12;
      }
      if (meridiem === 'A' && hour === 12) {
        hour = 0;
      }
      return {
        year: parseYear(dmyMatch[3]),
        month: parseInt(dmyMatch[2], 10),
        day: parseInt(dmyMatch[1], 10),
        hour: hour,
        minute: minute
      };
    }

    const parsed = new Date(value);
    if (!isNaN(parsed.getTime())) {
      return {
        year: parsed.getFullYear(),
        month: parsed.getMonth() + 1,
        day: parsed.getDate(),
        hour: parsed.getHours(),
        minute: parsed.getMinutes()
      };
    }

    return null;
  }

  function partsToDate(parts) {
    if (!parts) {
      return null;
    }
    const year = Number(parts.year);
    const month = Number(parts.month);
    const day = Number(parts.day);
    if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) {
      return null;
    }
    const hour = Number.isFinite(parts.hour) ? Number(parts.hour) : 0;
    const minute = Number.isFinite(parts.minute) ? Number(parts.minute) : 0;
    const second = Number.isFinite(parts.second) ? Number(parts.second) : 0;
    const date = new Date(year, month - 1, day, hour, minute, second, 0);
    if (isNaN(date.getTime())) {
      return null;
    }
    return date;
  }

  function parseDateValue(input) {
    if (input == null || input === '') {
      return null;
    }
    const parts = parseDateParts(input);
    if (!parts) {
      return null;
    }
    return partsToDate(parts);
  }

  function getDateSortValue(value) {
    const date = parseDateValue(value);
    if (!date) {
      return Number.POSITIVE_INFINITY;
    }
    const time = date.getTime();
    if (!Number.isFinite(time)) {
      return Number.POSITIVE_INFINITY;
    }
    return time;
  }

  function resolveLocale(locale) {
    if (typeof locale !== 'string' || !locale.trim()) {
      return DEFAULT_LOCALE;
    }
    return locale;
  }

  function fmtDate(value, locale) {
    const parts = parseDateParts(value);
    if (!parts) {
      if (value == null) return '';
      return String(value);
    }

    const normalizedLocale = resolveLocale(locale).toLowerCase();
    let separator = '/';
    let order = ['day', 'month', 'year'];

    if (normalizedLocale.startsWith('en')) {
      order = ['month', 'day', 'year'];
      separator = '/';
    } else if (normalizedLocale.startsWith('de')) {
      order = ['day', 'month', 'year'];
      separator = '.';
    } else if (normalizedLocale.includes('-')) {
      const language = normalizedLocale.split('-')[0];
      if (language === 'en') {
        order = ['month', 'day', 'year'];
        separator = '/';
      }
    }

    const dateSegments = order.map(function (key) {
      const valueKey = parts[key];
      if (key === 'year') {
        return String(valueKey).padStart(4, '0');
      }
      return pad2(valueKey);
    });

    const formattedDate = dateSegments.join(separator);
    const formattedTime = `${pad2(parts.hour)}:${pad2(parts.minute)}`;
    return `${formattedDate} ${formattedTime}`.trim();
  }

  function toDateInputValue(value) {
    const parts = parseDateParts(value);
    if (!parts) {
      return '';
    }
    const year = Number(parts.year);
    const month = Number(parts.month);
    const day = Number(parts.day);
    if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) {
      return '';
    }
    const hour = Number.isFinite(parts.hour) ? Number(parts.hour) : 0;
    const minute = Number.isFinite(parts.minute) ? Number(parts.minute) : 0;
    return (
      year +
      '-' + pad2(month) +
      '-' + pad2(day) +
      'T' + pad2(hour) +
      ':' + pad2(minute)
    );
  }

  function isValidDate(value) {
    return value instanceof Date && !isNaN(value.getTime());
  }

  function startOfDay(value) {
    if (!isValidDate(value)) {
      return null;
    }
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  function endOfDay(value) {
    if (!isValidDate(value)) {
      return null;
    }
    return new Date(value.getFullYear(), value.getMonth(), value.getDate(), 23, 59, 59, 999);
  }

  function isSameDay(a, b) {
    if (!isValidDate(a) || !isValidDate(b)) {
      return false;
    }
    return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
  }

  function normalizeDateRange(range) {
    const normalized = { start: null, end: null };

    if (range && range.start != null) {
      let startDate = range.start;
      if (!isValidDate(startDate)) {
        if (typeof startDate === 'string' && startDate.trim()) {
          startDate = parseDateValue(startDate);
        } else if (typeof startDate === 'number' && Number.isFinite(startDate)) {
          startDate = new Date(startDate);
        } else {
          startDate = null;
        }
      }
      if (isValidDate(startDate)) {
        normalized.start = startOfDay(startDate);
      }
    }

    if (range && range.end != null) {
      let endDate = range.end;
      if (!isValidDate(endDate)) {
        if (typeof endDate === 'string' && endDate.trim()) {
          endDate = parseDateValue(endDate);
        } else if (typeof endDate === 'number' && Number.isFinite(endDate)) {
          endDate = new Date(endDate);
        } else {
          endDate = null;
        }
      }
      if (isValidDate(endDate)) {
        normalized.end = endOfDay(endDate);
      }
    }

    if (normalized.start && normalized.end && normalized.start.getTime() > normalized.end.getTime()) {
      const originalStart = normalized.start;
      const originalEnd = normalized.end;
      normalized.start = startOfDay(originalEnd);
      normalized.end = endOfDay(originalStart);
    }

    return normalized;
  }

  function areRangesEqual(a, b) {
    const aStart = isValidDate(a && a.start) ? a.start.getTime() : null;
    const bStart = isValidDate(b && b.start) ? b.start.getTime() : null;
    const aEnd = isValidDate(a && a.end) ? a.end.getTime() : null;
    const bEnd = isValidDate(b && b.end) ? b.end.getTime() : null;
    return aStart === bStart && aEnd === bEnd;
  }

  function formatDateOnly(date, locale, includeYear) {
    if (!isValidDate(date)) {
      return '';
    }
    const options = includeYear
      ? { day: '2-digit', month: 'short', year: 'numeric' }
      : { day: '2-digit', month: 'short' };
    const formatter = new Intl.DateTimeFormat(resolveLocale(locale), options);
    return formatter.format(date);
  }

  function formatDateRangeLabel(range, locale) {
    const normalized = normalizeDateRange(range || {});
    const start = normalized.start;
    const end = normalized.end;

    if (!start && !end) {
      return 'Todos';
    }

    const today = new Date();
    if (start && end && isSameDay(start, end) && isSameDay(start, today)) {
      return 'Hoy';
    }

    const yesterday = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1);
    if (start && end && isSameDay(start, end) && isSameDay(start, yesterday)) {
      return 'Ayer';
    }

    if (start && end && isSameDay(start, end)) {
      return formatDateOnly(start, locale, true);
    }

    if (start && end) {
      const sameYear = start.getFullYear() === end.getFullYear();
      const startLabel = formatDateOnly(start, locale, !sameYear);
      const endLabel = formatDateOnly(end, locale, true);
      return `${startLabel} – ${endLabel}`;
    }

    if (start) {
      return `Desde ${formatDateOnly(start, locale, true)}`;
    }

    if (end) {
      return `Hasta ${formatDateOnly(end, locale, true)}`;
    }

    return 'Todos';
  }

  function dateToInputDateValue(date) {
    if (!isValidDate(date)) {
      return '';
    }
    const year = date.getFullYear();
    const month = pad2(date.getMonth() + 1);
    const day = pad2(date.getDate());
    return `${year}-${month}-${day}`;
  }

  function parseInputDateValue(value) {
    if (!value) {
      return null;
    }
    const parts = parseDateParts(value);
    if (!parts) {
      return null;
    }
    const year = Number(parts.year);
    const month = Number(parts.month);
    const day = Number(parts.day);
    if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) {
      return null;
    }
    return new Date(year, month - 1, day);
  }

  function createTodayRange() {
    const today = new Date();
    return normalizeDateRange({ start: today, end: today });
  }

  function toApiDateValue(value) {
    if (value == null || value === '') {
      return '';
    }
    const parts = parseDateParts(value);
    if (!parts) {
      return '';
    }
    const year = Number(parts.year);
    const month = Number(parts.month);
    const day = Number(parts.day);
    if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) {
      return '';
    }
    const hour = Number.isFinite(parts.hour) ? Number(parts.hour) : 0;
    const minute = Number.isFinite(parts.minute) ? Number(parts.minute) : 0;
    const second = Number.isFinite(parts.second) ? Number(parts.second) : 0;
    return (
      year +
      '-' + pad2(month) +
      '-' + pad2(day) +
      'T' + pad2(hour) +
      ':' + pad2(minute) +
      ':' + pad2(second)
    );
  }

  function getColumnKeyFromHeader(headerLabel) {
    if (headerLabel == null) {
      return null;
    }
    const normalized = String(headerLabel).trim().toLowerCase();
    if (!normalized) {
      return null;
    }
    if (Object.prototype.hasOwnProperty.call(COLUMN_LABEL_TO_KEY, normalized)) {
      return COLUMN_LABEL_TO_KEY[normalized];
    }
    return null;
  }

  const DAILY_VIEW_ALLOWED_STATUS_SET = new Set(
    ['drop', 'live', 'qro yard', 'mty yard', 'loading', 'in transit mx'].map(function (status) {
      return status.toLowerCase();
    })
  );

  const TODAY_DELIVERY_EXCLUDED_STATUS_SET = new Set(
    ['at destination', 'delivered', 'cancelled'].map(function (status) {
      return status.toLowerCase();
    })
  );

  const CONFIRMED_APPOINTMENTS_EXCLUDED_STATUS_SET = new Set(
    ['delivered', 'live', 'drop', 'cancelled'].map(function (status) {
      return status.toLowerCase();
    })
  );

  const CRUCES_EXCLUDED_STATUS_SET = new Set(
    ['cancelled', 'in transit usa', 'at destination', 'delivered'].map(function (status) {
      return status.toLowerCase();
    })
  );

  function matchesDailyLoadsView(row, context) {
    if (!row || !context || !context.columnMap) {
      return false;
    }
    const columnMap = context.columnMap;
    const citaCargaIndex = columnMap.citaCarga;
    const estatusIndex = columnMap.estatus;
    if (citaCargaIndex == null || estatusIndex == null) {
      return false;
    }
    const rawStatus = row[estatusIndex];
    const rawDate = row[citaCargaIndex];
    if (rawStatus == null || rawStatus === '' || rawDate == null || rawDate === '') {
      return false;
    }
    const normalizedStatus = String(rawStatus).trim().toLowerCase();
    if (!DAILY_VIEW_ALLOWED_STATUS_SET.has(normalizedStatus)) {
      return false;
    }
    const citaDate = parseDateValue(rawDate);
    if (!citaDate) {
      return false;
    }
    const now = context.now instanceof Date ? context.now : new Date();
    const startOfTomorrow = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);
    return citaDate < startOfTomorrow;
  }

  function matchesInventarioNlarView(row, context) {
    if (!row || !context || !context.columnMap) {
      return false;
    }
    const columnMap = context.columnMap;
    const estatusIndex = columnMap.estatus;
    if (estatusIndex == null) {
      return false;
    }
    const rawStatus = row[estatusIndex];
    if (rawStatus == null || rawStatus === '') {
      return false;
    }
    const normalizedStatus = String(rawStatus).trim().toLowerCase();
    return normalizedStatus === 'nuevo laredo yard';
  }

  function matchesWeeklyProgramView(row, context) {
    if (!row || !context || !context.columnMap) {
      return false;
    }
    const columnMap = context.columnMap;
    const citaCargaIndex = columnMap.citaCarga;
    if (citaCargaIndex == null) {
      return false;
    }
    const rawDate = row[citaCargaIndex];
    if (rawDate == null || rawDate === '') {
      return false;
    }
    const citaDate = parseDateValue(rawDate);
    if (!citaDate) {
      return false;
    }
    const reference =
      context.now instanceof Date && !isNaN(context.now.getTime()) ? new Date(context.now.getTime()) : new Date();
    const startOfDay = new Date(reference.getFullYear(), reference.getMonth(), reference.getDate());
    const day = startOfDay.getDay();
    const diffToMonday = day === 0 ? -6 : 1 - day;
    startOfDay.setDate(startOfDay.getDate() + diffToMonday);
    startOfDay.setHours(0, 0, 0, 0);
    const endOfWeek = new Date(startOfDay.getTime());
    endOfWeek.setDate(endOfWeek.getDate() + 7);
    return citaDate >= startOfDay && citaDate < endOfWeek;
  }

  function matchesTodayDeliveriesView(row, context) {
    if (!row || !context || !context.columnMap) {
      return false;
    }
    const columnMap = context.columnMap;
    const estatusIndex = columnMap.estatus;
    const citaEntregaIndex = columnMap.citaEntrega;
    const llegadaEntregaIndex = columnMap.llegadaEntrega;
    if (estatusIndex == null || citaEntregaIndex == null) {
      return false;
    }

    let rawStatus = null;
    if (estatusIndex < row.length) {
      rawStatus = row[estatusIndex];
    }
    if (rawStatus != null && rawStatus !== '') {
      const normalizedStatus = String(rawStatus).trim().toLowerCase();
      if (TODAY_DELIVERY_EXCLUDED_STATUS_SET.has(normalizedStatus)) {
        return false;
      }
    }

    const now = context.now instanceof Date && !isNaN(context.now.getTime()) ? context.now : new Date();
    const startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const startOfTomorrow = new Date(startOfToday.getTime());
    startOfTomorrow.setDate(startOfTomorrow.getDate() + 1);

    let hasCitaEntrega = false;
    let isDue = false;

    if (citaEntregaIndex != null && citaEntregaIndex >= 0 && citaEntregaIndex < row.length) {
      const rawCita = row[citaEntregaIndex];
      if (rawCita != null && rawCita !== '') {
        const citaDate = parseDateValue(rawCita);
        if (citaDate) {
          hasCitaEntrega = true;
          if (citaDate < startOfTomorrow) {
            isDue = true;
          }
        }
      }
    }

    if (!isDue && llegadaEntregaIndex != null && llegadaEntregaIndex >= 0 && llegadaEntregaIndex < row.length) {
      const rawLlegada = row[llegadaEntregaIndex];
      if (rawLlegada != null && rawLlegada !== '') {
        const llegadaDate = parseDateValue(rawLlegada);
        if (llegadaDate && llegadaDate < startOfTomorrow) {
          isDue = true;
        }
      }
    }

    return hasCitaEntrega && isDue;
  }

  function matchesConfirmedAppointmentsView(row, context) {
    if (!row || !context || !context.columnMap) {
      return false;
    }
    const columnMap = context.columnMap;
    const citaEntregaIndex = columnMap.citaEntrega;
    if (citaEntregaIndex == null) {
      return false;
    }

    let rawCitaEntrega = null;
    if (citaEntregaIndex >= 0 && citaEntregaIndex < row.length) {
      rawCitaEntrega = row[citaEntregaIndex];
    }
    if (rawCitaEntrega == null || rawCitaEntrega === '') {
      return false;
    }

    const estatusIndex = columnMap.estatus;
    if (estatusIndex != null && estatusIndex >= 0 && estatusIndex < row.length) {
      const rawStatus = row[estatusIndex];
      if (rawStatus != null && rawStatus !== '') {
        const normalizedStatus = String(rawStatus).trim().toLowerCase();
        if (CONFIRMED_APPOINTMENTS_EXCLUDED_STATUS_SET.has(normalizedStatus)) {
          return false;
        }
      }
    }

    return true;
  }

  function matchesCrucesView(row, context) {
    if (!row || !context || !context.columnMap) {
      return false;
    }
    const columnMap = context.columnMap;
    const segmentoIndex = columnMap.segmento;
    if (segmentoIndex == null) {
      return false;
    }

    const segmentValue = segmentoIndex >= 0 && segmentoIndex < row.length ? row[segmentoIndex] : null;
    if (segmentValue == null || segmentValue === '') {
      return false;
    }

    const normalizedSegment = String(segmentValue).trim().toLowerCase();
    if (normalizedSegment !== 'otr' && normalizedSegment !== 'reg') {
      return false;
    }

    const estatusIndex = columnMap.estatus;
    if (estatusIndex != null && estatusIndex >= 0 && estatusIndex < row.length) {
      const rawStatus = row[estatusIndex];
      if (rawStatus != null && rawStatus !== '') {
        const normalizedStatus = String(rawStatus).trim().toLowerCase();
        if (CRUCES_EXCLUDED_STATUS_SET.has(normalizedStatus)) {
          return false;
        }
      }
    }

    if (typeof columnMap.citaEntrega !== 'number') {
      return false;
    }

    const citaEntregaIndex = columnMap.citaEntrega;
    if (citaEntregaIndex < 0 || citaEntregaIndex >= row.length) {
      return false;
    }

    const rawCitaEntrega = row[citaEntregaIndex];
    if (rawCitaEntrega == null || rawCitaEntrega === '') {
      return false;
    }

    const citaDate = parseDateValue(rawCitaEntrega);
    if (!citaDate) {
      return false;
    }

    const now = context.now instanceof Date && !isNaN(context.now.getTime()) ? context.now : new Date();
    const startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const startOfCitaEntrega = new Date(
      citaDate.getFullYear(),
      citaDate.getMonth(),
      citaDate.getDate()
    );
    const diffInDays = Math.floor((startOfCitaEntrega.getTime() - startOfToday.getTime()) / MS_PER_DAY);

    if (normalizedSegment === 'otr') {
      return diffInDays <= 3;
    }

    if (normalizedSegment === 'reg') {
      return diffInDays <= 1;
    }

    return false;
  }

  const TABLE_VIEWS = [
    {
      id: 'all',
      label: 'Todas las cargas',
      filter: function () {
        return true;
      },
      dateFilterEnabled: false
    },
    {
      id: 'daily-loads',
      label: 'Cargas diarias',
      filter: matchesDailyLoadsView
    },
    {
      id: 'cruces',
      label: 'Cruces',
      filter: matchesCrucesView
    },
    {
      id: 'today-deliveries',
      label: 'Entregas hoy',
      filter: matchesTodayDeliveriesView
    },
    {
      id: 'confirmed-appointments',
      label: 'Citas confirmadas',
      filter: matchesConfirmedAppointmentsView
    },
    {
      id: 'inventario-nlar',
      label: 'Inventario Nlar',
      filter: matchesInventarioNlarView
    },
    {
      id: 'weekly-program',
      label: 'Programa semanal',
      filter: matchesWeeklyProgramView
    }
  ];

  function columnLetter(index) {
    let result = '';
    let i = index;
    while (i >= 0) {
      result = String.fromCharCode((i % 26) + 65) + result;
      i = Math.floor(i / 26) - 1;
    }
    return result;
  }

  function isDateHeader(label) {
    if (label == null) {
      return false;
    }
    const text = String(label).trim();
    if (!text) {
      return false;
    }
    return DATE_HEADER_REGEX.test(text);
  }

  function formatHeaderLabel(label) {
    if (label == null) {
      return '';
    }
    const text = String(label).trim();
    if (!text) {
      return '';
    }
    const upper = text.toUpperCase();
    const lower = text.toLowerCase();
    const hasLetters = upper !== lower;
    const isAllCaps = hasLetters && text === upper;
    if (isAllCaps) {
      return text;
    }
    return text.charAt(0).toUpperCase() + text.slice(1);
  }

  function getStoredValue(key) {
    try {
      if (global.localStorage) {
        return global.localStorage.getItem(key);
      }
    } catch (err) {
      return null;
    }
    return null;
  }

  function setStoredValue(key, value) {
    try {
      if (!global.localStorage) {
        return;
      }
      if (value === undefined || value === null) {
        global.localStorage.removeItem(key);
        return;
      }
      const payload = typeof value === 'string' ? value : JSON.stringify(value);
      global.localStorage.setItem(key, payload);
    } catch (err) {
      // Ignore storage errors (private mode, etc.)
    }
  }

  function normalizeUsers(config) {
    const result = [];
    if (!config) {
      return result;
    }

    function pushUser(raw) {
      if (!raw) return;
      const username = raw.usuario || raw.username || raw.user || raw.email;
      const password = raw.password || raw.pass || raw.clave;
      if (!username || !password) return;
      const displayName = raw.nombre || raw.displayName || raw.name || username;
      result.push({
        id: String(username).trim().toLowerCase(),
        username: String(username).trim(),
        password: String(password),
        displayName: String(displayName)
      });
    }

    if (Array.isArray(config.users)) {
      config.users.forEach(pushUser);
    }

    if (config.usuario && config.password) {
      pushUser(config);
    }

    const seen = new Set();
    return result.filter(function (user) {
      if (!user) return false;
      if (seen.has(user.id)) return false;
      seen.add(user.id);
      return true;
    });
  }

  function loadStoredUser() {
    const raw = getStoredValue(STORAGE_USER_KEY);
    if (!raw) {
      return null;
    }
    try {
      return JSON.parse(raw);
    } catch (err) {
      return null;
    }
  }

  async function fetchSecureConfig(url) {
    if (!url) {
      return {};
    }
    try {
      const response = await fetch(url, { cache: 'no-store' });
      if (!response.ok) {
        throw new Error('Secure config not available');
      }
      return await response.json();
    } catch (err) {
      return {};
    }
  }

  function normalizeObjectRows(rows) {
    if (!Array.isArray(rows) || rows.length === 0) {
      return [];
    }

    const configuredColumns = COLUMN_CONFIG.filter(function (column) {
      return rows.some(function (row) {
        return row && Object.prototype.hasOwnProperty.call(row, column.key);
      });
    });

    const knownKeys = configuredColumns.map(function (column) {
      return column.key;
    });

    const extraKeys = [];
    rows.forEach(function (row) {
      if (!row || typeof row !== 'object') {
        return;
      }
      Object.keys(row).forEach(function (key) {
        if (knownKeys.indexOf(key) === -1 && extraKeys.indexOf(key) === -1) {
          extraKeys.push(key);
        }
      });
    });

    const headers = configuredColumns
      .map(function (column) { return column.label; })
      .concat(extraKeys);

    const values = rows.map(function (row) {
      const baseValues = configuredColumns.map(function (column) {
        return row && Object.prototype.hasOwnProperty.call(row, column.key)
          ? row[column.key]
          : '';
      });
      const extraValues = extraKeys.map(function (key) {
        return row && Object.prototype.hasOwnProperty.call(row, key)
          ? row[key]
          : '';
      });
      return baseValues.concat(extraValues);
    });

    return [headers].concat(values);
  }

  async function fetchSheetData(apiBase, token) {
    if (!apiBase) {
      throw new Error('Falta configurar la URL del Apps Script.');
    }
    const url = new URL(apiBase);
    if (token) {
      url.searchParams.set('token', token);
    }
    url.searchParams.set('t', Date.now().toString());
    try {
      const response = await fetch(url.toString(), {
        headers: { Accept: 'application/json' },
        cache: 'no-store'
      });
      const payload = await response.json().catch(function () { return {}; });
      if (!response.ok) {
        const error = payload && payload.error ? payload.error : `Error ${response.status}`;
        const err = new Error(error);
        err.status = response.status;
        throw err;
      }
      if (!payload || !Array.isArray(payload.data)) {
        return [];
      }
      if (payload.data.length === 0) {
        return [];
      }

      const firstRow = payload.data[0];
      if (Array.isArray(firstRow)) {
        return payload.data;
      }

      if (firstRow && typeof firstRow === 'object') {
        return normalizeObjectRows(payload.data);
      }

      return [];
    } catch (err) {
      if (err instanceof Error) {
        if (!err.message || err.message === 'Failed to fetch') {
          const friendly = new Error('No se pudo conectar con el Apps Script.');
          friendly.cause = err;
          throw friendly;
        }
        throw err;
      }
      throw new Error('No se pudo conectar con el Apps Script.');
    }
  }

  async function submitRecordRequest(apiBase, token, payload) {
    if (!apiBase) {
      throw new Error('Falta configurar la URL del Apps Script.');
    }
    const url = new URL(apiBase);
    if (token) {
      url.searchParams.set('token', token);
    }
    const params = new URLSearchParams();
    Object.keys(payload || {}).forEach(function (key) {
      const value = payload[key];
      if (value == null) {
        params.append(key, '');
      } else {
        params.append(key, String(value));
      }
    });
    try {
      const response = await fetch(url.toString(), {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8',
          Accept: 'application/json'
        },
        body: params.toString()
      });
      const result = await response.json().catch(function () { return {}; });
      if (!response.ok) {
        const errorMessage = result && result.error ? result.error : `Error ${response.status}`;
        const err = new Error(errorMessage);
        err.status = response.status;
        throw err;
      }
      if (result && result.error) {
        const err = new Error(result.error);
        err.status = response.status;
        throw err;
      }
      return result;
    } catch (err) {
      if (err instanceof Error) {
        if (!err.message || err.message === 'Failed to fetch') {
          const friendly = new Error('No se pudo conectar con el Apps Script.');
          friendly.cause = err;
          throw friendly;
        }
        throw err;
      }
      throw new Error('No se pudo conectar con el Apps Script.');
    }
  }

  function initApp() {
    if (!global.document) {
      return;
    }

    const doc = global.document;
    const appRoot = doc.querySelector('[data-app]');
    if (!appRoot) {
      return;
    }

    const refs = {
      tableHead: doc.querySelector('[data-table-head]'),
      tableBody: doc.querySelector('[data-table-body]'),
      viewMenu: doc.querySelector('[data-view-menu]'),
      status: doc.querySelector('[data-status]'),
      refreshButton: doc.querySelector('[data-action="refresh"]'),
      newRecordButton: doc.querySelector('[data-action="new-record"]'),
      logoutButton: doc.querySelector('[data-action="logout"]'),
      changeTokenButton: doc.querySelector('[data-action="change-token"]'),
      themeSwitch: doc.querySelector('[data-theme-switch]'),
      themeLabel: doc.querySelector('[data-theme-label]'),
      filterSearchInput: doc.querySelector('[data-filter-search]'),
      dateFilter: doc.querySelector('[data-date-filter]'),
      dateLabel: doc.querySelector('[data-date-label]'),
      datePopover: doc.querySelector('[data-date-popover]'),
      dateStartInput: doc.querySelector('[data-date-start]'),
      dateEndInput: doc.querySelector('[data-date-end]'),
      datePrevButton: doc.querySelector('[data-action="date-prev"]'),
      dateNextButton: doc.querySelector('[data-action="date-next"]'),
      dateToggleButton: doc.querySelector('[data-action="date-toggle"]'),
      dateClearButton: doc.querySelector('[data-action="date-clear"]'),
      statusFilter: doc.querySelector('[data-status-filter]'),
      statusToggleButton: doc.querySelector('[data-action="status-toggle"]'),
      statusLabel: doc.querySelector('[data-status-label]'),
      statusPopover: doc.querySelector('[data-status-popover]'),
      statusList: doc.querySelector('[data-status-list]'),
      lastUpdated: doc.querySelector('[data-last-updated]'),
      currentUser: doc.querySelector('[data-current-user]'),
      loginModal: doc.querySelector('[data-login-modal]'),
      loginForm: doc.querySelector('[data-login-form]'),
      loginError: doc.querySelector('[data-login-error]'),
      loginTokenField: doc.querySelector('[data-token-field]'),
      backdrop: doc.querySelector('[data-backdrop]'),
      editModal: doc.querySelector('[data-edit-modal]'),
      editForm: doc.querySelector('[data-edit-form]'),
      editError: doc.querySelector('[data-edit-error]'),
      cancelEditButton: doc.querySelector('[data-action="cancel-edit"]'),
      editTitle: doc.querySelector('[data-edit-title]'),
      editHint: doc.querySelector('[data-edit-hint]'),
      editSubmitButton: doc.querySelector('[data-edit-submit]')
    };

    const rawStoredTheme = getStoredValue(STORAGE_THEME_KEY);
    const normalizedStoredTheme = rawStoredTheme ? String(rawStoredTheme).toLowerCase() : '';
    const hasInitialStoredTheme = normalizedStoredTheme === THEME_DARK || normalizedStoredTheme === THEME_LIGHT;
    const prefersDarkMedia = typeof global.matchMedia === 'function' ? global.matchMedia('(prefers-color-scheme: dark)') : null;
    let hasStoredTheme = hasInitialStoredTheme;
    const initialTheme = hasInitialStoredTheme
      ? normalizedStoredTheme
      : prefersDarkMedia && prefersDarkMedia.matches
        ? THEME_DARK
        : THEME_LIGHT;

    const state = {
      config: global.APP_CONFIG || { API_BASE: '', SECURE_CONFIG_URL: '' },
      token: '',
      users: [],
      data: [],
      locale: (global.navigator && global.navigator.language) || DEFAULT_LOCALE,
      currentUser: null,
      loading: false,
      secureConfigLoaded: false,
      editingRecord: null,
      currentViewId: TABLE_VIEWS[0] ? TABLE_VIEWS[0].id : 'all',
      theme: initialTheme,
      filters: {
        searchText: '',
        dateRange: null,
        status: ''
      },
      availableStatuses: [],
      isDatePopoverOpen: false,
      isStatusPopoverOpen: false
    };

    function setTheme(theme, options) {
      const normalized = theme === THEME_DARK ? THEME_DARK : THEME_LIGHT;
      state.theme = normalized;
      doc.documentElement.setAttribute('data-theme', normalized);
      if (refs.themeSwitch) {
        refs.themeSwitch.checked = normalized === THEME_DARK;
        const ariaLabel = normalized === THEME_DARK ? 'Cambiar a tema claro' : 'Cambiar a tema oscuro';
        refs.themeSwitch.setAttribute('aria-label', ariaLabel);
      }
      if (refs.themeLabel) {
        refs.themeLabel.textContent = normalized === THEME_DARK ? 'Tema oscuro' : 'Tema claro';
      }
      if (!options || options.persist !== false) {
        setStoredValue(STORAGE_THEME_KEY, normalized);
        hasStoredTheme = true;
      }
      return normalized;
    }

    function handleThemeSwitchChange(event) {
      const target = event && event.target;
      const isChecked = target ? Boolean(target.checked) : false;
      setTheme(isChecked ? THEME_DARK : THEME_LIGHT);
    }

    function handleSystemThemeChange(event) {
      if (hasStoredTheme) {
        return;
      }
      const nextTheme = event && event.matches ? THEME_DARK : THEME_LIGHT;
      setTheme(nextTheme, { persist: false });
    }

    setTheme(initialTheme, { persist: false });

    if (prefersDarkMedia) {
      if (typeof prefersDarkMedia.addEventListener === 'function') {
        prefersDarkMedia.addEventListener('change', handleSystemThemeChange);
      } else if (typeof prefersDarkMedia.addListener === 'function') {
        prefersDarkMedia.addListener(handleSystemThemeChange);
      }
    }

    let wasDatePopoverOpen = false;
    let wasStatusPopoverOpen = false;

    if (refs.filterSearchInput) {
      refs.filterSearchInput.value = state.filters.searchText;
    }

    renderDateFilter();
    renderStatusFilter();

    const EDIT_MODAL_CONTENT = {
      edit: {
        title: 'Editar registro',
        hint: 'Actualiza la información del viaje y guarda los cambios.',
        submit: 'Guardar cambios'
      },
      create: {
        title: 'Nuevo registro',
        hint: 'Captura la información del viaje y guarda el registro.',
        submit: 'Agregar registro'
      }
    };


    const BULK_ALLOWED_EXTENSIONS = new Set(['xlsx', 'xlsm']);
    const BULK_REQUIRED_HEADERS = ['Trip', 'Ejecutivo'];
    const BULK_OPTIONAL_HEADERS = [
      'Caja',
      'Estatus',
      'Cliente',
      'Destino',
      'Referencia',
      'Segmento',
      'Cita carga',
      'Cita entrega'
    ];
    const BULK_CANONICAL_HEADERS = BULK_REQUIRED_HEADERS.concat(BULK_OPTIONAL_HEADERS);
    const BULK_HEADER_NORMALIZATION_MAP = (function () {
      const map = {};
      BULK_CANONICAL_HEADERS.forEach(function (label) {
        if (label) {
          map[label.trim().toLowerCase()] = label;
        }
      });
      map['status'] = 'Estatus';
      map['cita carga'] = 'Cita carga';
      map['citacarga'] = 'Cita carga';
      map['cita de carga'] = 'Cita carga';
      map['cita_carga'] = 'Cita carga';
      map['cita entrega'] = 'Cita entrega';
      map['citaentrega'] = 'Cita entrega';
      map['cita de entrega'] = 'Cita entrega';
      map['cita_entrega'] = 'Cita entrega';
      return map;
    })();
    const BULK_DATE_HEADER_SET = new Set(['Cita carga', 'Cita entrega']);
    const BUILTIN_DATE_FORMAT_IDS = new Set([
      14, 15, 16, 17, 18, 19, 20, 21, 22,
      27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
      45, 46, 47,
      50, 51, 52, 53, 54, 55, 56, 57, 58
    ]);

    function normalizeBulkHeader(name) {
      if (name == null) {
        return '';
      }
      const normalized = String(name).trim().toLowerCase();
      if (!normalized) {
        return '';
      }
      if (Object.prototype.hasOwnProperty.call(BULK_HEADER_NORMALIZATION_MAP, normalized)) {
        return BULK_HEADER_NORMALIZATION_MAP[normalized];
      }
      return '';
    }

    function getFileExtension(filename) {
      if (!filename) {
        return '';
      }
      const match = String(filename)
        .trim()
        .toLowerCase()
        .match(/\.([0-9a-z]+)$/i);
      return match ? match[1] : '';
    }

    function normalizeZipPath(path) {
      if (!path) {
        return '';
      }
      const segments = [];
      const parts = String(path)
        .replace(/\\/g, '/')
        .split('/');
      parts.forEach(function (part) {
        if (!part || part === '.') {
          return;
        }
        if (part === '..') {
          if (segments.length > 0) {
            segments.pop();
          }
          return;
        }
        segments.push(part);
      });
      return segments.join('/');
    }

    function resolveSheetPath(target) {
      if (!target) {
        return '';
      }
      let normalized = String(target).trim();
      if (!normalized) {
        return '';
      }
      if (normalized[0] === '/') {
        normalized = normalized.slice(1);
      }
      if (!/^xl\//i.test(normalized)) {
        normalized = `xl/${normalized}`;
      }
      return normalizeZipPath(normalized);
    }

    function parseXmlDocument(xmlText) {
      if (typeof global.DOMParser !== 'function') {
        throw new Error('Este navegador no soporta la lectura de archivos de Excel.');
      }
      const parser = new global.DOMParser();
      const doc = parser.parseFromString(String(xmlText || ''), 'application/xml');
      const errorNode = doc.getElementsByTagName('parsererror')[0];
      if (errorNode) {
        throw new Error('El archivo de Excel contiene datos XML no válidos.');
      }
      return doc;
    }

    function excelSerialToDate(serial) {
      if (!Number.isFinite(serial)) {
        return null;
      }
      const adjusted = serial >= 60 ? serial - 1 : serial;
      const wholeDays = Math.floor(adjusted);
      const fractionalDay = adjusted - wholeDays;
      const baseDate = new Date(1899, 11, 30);
      baseDate.setDate(baseDate.getDate() + wholeDays);
      if (fractionalDay > 0) {
        const totalSeconds = Math.round(fractionalDay * 24 * 60 * 60);
        let seconds = totalSeconds % 60;
        let totalMinutes = (totalSeconds - seconds) / 60;
        let minutes = totalMinutes % 60;
        let hours = (totalMinutes - minutes) / 60;
        if (hours >= 24) {
          baseDate.setDate(baseDate.getDate() + Math.floor(hours / 24));
          hours = hours % 24;
        }
        baseDate.setHours(hours, minutes, seconds, 0);
      } else {
        baseDate.setHours(0, 0, 0, 0);
      }
      return baseDate;
    }

    function formatDateForApi(date) {
      if (!(date instanceof Date) || isNaN(date.getTime())) {
        return '';
      }
      return (
        date.getFullYear() +
        '-' + pad2(date.getMonth() + 1) +
        '-' + pad2(date.getDate()) +
        'T' + pad2(date.getHours()) +
        ':' + pad2(date.getMinutes()) +
        ':' + pad2(date.getSeconds())
      );
    }

    function convertBulkDateValue(input) {
      if (input == null || input === '') {
        return { value: '' };
      }
      if (input instanceof Date) {
        if (isNaN(input.getTime())) {
          return { value: '', error: 'Fecha inválida' };
        }
        return { value: formatDateForApi(input) };
      }
      if (typeof input === 'number') {
        const date = excelSerialToDate(input);
        if (!date) {
          return { value: '', error: 'Fecha inválida' };
        }
        return { value: formatDateForApi(date) };
      }
      const trimmed = String(input).trim();
      if (!trimmed) {
        return { value: '' };
      }
      if (/^\d+(?:\.\d+)?$/.test(trimmed)) {
        const numericValue = Number(trimmed);
        if (Number.isFinite(numericValue)) {
          const numericDate = excelSerialToDate(numericValue);
          if (numericDate) {
            return { value: formatDateForApi(numericDate) };
          }
        }
      }
      const isoValue = toApiDateValue(trimmed);
      if (!isoValue) {
        return { value: '', error: 'Formato de fecha no reconocido' };
      }
      return { value: isoValue };
    }

    function summarizeValues(values, maxItems) {
      if (!Array.isArray(values) || values.length === 0) {
        return '';
      }
      const limit = typeof maxItems === 'number' && maxItems > 0 ? maxItems : values.length;
      const visible = values.slice(0, limit);
      const remaining = values.length - visible.length;
      const summary = visible.join(', ');
      if (remaining > 0) {
        return `${summary} y ${remaining} más`;
      }
      return summary;
    }

    function summarizeIssues(issues, maxItems) {
      if (!Array.isArray(issues) || issues.length === 0) {
        return '';
      }
      const limit = typeof maxItems === 'number' && maxItems > 0 ? maxItems : issues.length;
      const visible = issues.slice(0, limit);
      let summary = visible.join(' ');
      const remaining = issues.length - visible.length;
      if (remaining > 0) {
        summary += ` (+${remaining} filas adicionales)`;
      }
      return summary;
    }

    function columnLabelToIndex(cellReference) {
      if (!cellReference) {
        return null;
      }
      const match = String(cellReference).match(/^[A-Z]+/i);
      if (!match) {
        return null;
      }
      const letters = match[0].toUpperCase();
      let index = 0;
      for (let i = 0; i < letters.length; i++) {
        index = index * 26 + (letters.charCodeAt(i) - 64);
      }
      return index - 1;
    }

    function parseSharedStringsXml(xmlText) {
      const doc = parseXmlDocument(xmlText);
      const siNodes = doc.getElementsByTagName('si');
      const strings = [];
      for (let i = 0; i < siNodes.length; i++) {
        const si = siNodes[i];
        const tNodes = si.getElementsByTagName('t');
        if (tNodes.length === 0) {
          strings.push('');
          continue;
        }
        let text = '';
        for (let j = 0; j < tNodes.length; j++) {
          text += tNodes[j].textContent || '';
        }
        strings.push(text);
      }
      return strings;
    }

    function isDateFormatCode(formatCode) {
      if (!formatCode) {
        return false;
      }
      const cleaned = String(formatCode)
        .replace(/"[^"]*"/g, '')
        .replace(/\[[^\]]*\]/g, '')
        .toLowerCase();
      if (!cleaned) {
        return false;
      }
      if (cleaned.includes('am/pm')) {
        return true;
      }
      const hasYear = cleaned.indexOf('y') > -1;
      const hasDay = cleaned.indexOf('d') > -1;
      const hasMonth = cleaned.indexOf('m') > -1;
      const hasHour = cleaned.indexOf('h') > -1;
      const hasSecond = cleaned.indexOf('s') > -1;
      if ((hasYear && hasMonth) || (hasDay && hasMonth) || (hasYear && hasDay)) {
        return true;
      }
      if (hasHour && (hasMonth || hasDay || hasSecond)) {
        return true;
      }
      if (hasHour && cleaned.indexOf('m') > -1) {
        return true;
      }
      if (hasSecond && cleaned.indexOf('m') > -1) {
        return true;
      }
      return false;
    }

    function parseStylesXml(xmlText) {
      const doc = parseXmlDocument(xmlText);
      const numFmtNodes = doc.getElementsByTagName('numFmt');
      const numFmtMap = {};
      for (let i = 0; i < numFmtNodes.length; i++) {
        const node = numFmtNodes[i];
        const idAttr = node.getAttribute('numFmtId');
        const codeAttr = node.getAttribute('formatCode');
        if (idAttr == null) {
          continue;
        }
        const numFmtId = parseInt(idAttr, 10);
        if (Number.isFinite(numFmtId)) {
          numFmtMap[numFmtId] = codeAttr || '';
        }
      }
      const cellXfs = [];
      const cellXfsNode = doc.getElementsByTagName('cellXfs')[0];
      if (cellXfsNode) {
        const xfNodes = cellXfsNode.getElementsByTagName('xf');
        for (let i = 0; i < xfNodes.length; i++) {
          const xf = xfNodes[i];
          const numFmtIdAttr = xf.getAttribute('numFmtId');
          const numFmtId = numFmtIdAttr == null ? NaN : parseInt(numFmtIdAttr, 10);
          cellXfs.push(Number.isFinite(numFmtId) ? numFmtId : null);
        }
      }
      return { numFmtMap: numFmtMap, cellXfs: cellXfs };
    }

    function isDateStyle(styleIndex, stylesInfo) {
      if (!stylesInfo || !Array.isArray(stylesInfo.cellXfs)) {
        return false;
      }
      if (!Number.isFinite(styleIndex) || styleIndex < 0 || styleIndex >= stylesInfo.cellXfs.length) {
        return false;
      }
      const numFmtId = stylesInfo.cellXfs[styleIndex];
      if (!Number.isFinite(numFmtId)) {
        return false;
      }
      if (BUILTIN_DATE_FORMAT_IDS.has(numFmtId)) {
        return true;
      }
      if (!stylesInfo.numFmtMap || !Object.prototype.hasOwnProperty.call(stylesInfo.numFmtMap, numFmtId)) {
        return false;
      }
      return isDateFormatCode(stylesInfo.numFmtMap[numFmtId]);
    }

    function readCellValue(cellNode, sharedStrings, stylesInfo) {
      if (!cellNode) {
        return '';
      }
      const typeAttr = cellNode.getAttribute('t') || '';
      const styleAttr = cellNode.getAttribute('s');
      const styleIndex = styleAttr == null ? NaN : parseInt(styleAttr, 10);

      if (typeAttr === 'inlineStr') {
        const tNodes = cellNode.getElementsByTagName('t');
        if (tNodes.length === 0) {
          return '';
        }
        let text = '';
        for (let i = 0; i < tNodes.length; i++) {
          text += tNodes[i].textContent || '';
        }
        return text;
      }

      if (typeAttr === 'b') {
        const valueNode = cellNode.getElementsByTagName('v')[0];
        const raw = valueNode ? valueNode.textContent || '' : '';
        return raw === '1' ? '1' : raw === '0' ? '0' : raw;
      }

      const valueNode = cellNode.getElementsByTagName('v')[0];
      const rawValue = valueNode ? valueNode.textContent || '' : '';

      if (typeAttr === 's') {
        const index = parseInt(rawValue, 10);
        if (Number.isFinite(index) && Array.isArray(sharedStrings) && index >= 0 && index < sharedStrings.length) {
          return sharedStrings[index];
        }
        return '';
      }

      if (typeAttr === 'str') {
        return rawValue;
      }

      if (typeAttr === 'd') {
        const dateValue = new Date(rawValue);
        if (!isNaN(dateValue.getTime())) {
          return dateValue;
        }
        return rawValue;
      }

      if (!rawValue) {
        return '';
      }

      const numeric = Number(rawValue);
      if (Number.isFinite(numeric)) {
        if (isDateStyle(styleIndex, stylesInfo)) {
          const date = excelSerialToDate(numeric);
          return date || '';
        }
        return numeric;
      }

      return rawValue;
    }

    function extractSheetRows(sheetXml, sharedStrings, stylesInfo) {
      const doc = parseXmlDocument(sheetXml);
      const sheetData = doc.getElementsByTagName('sheetData')[0];
      if (!sheetData) {
        return [];
      }
      const rowNodes = sheetData.getElementsByTagName('row');
      const rows = [];
      let maxColumnIndex = 0;
      for (let i = 0; i < rowNodes.length; i++) {
        const rowNode = rowNodes[i];
        const rowIndexAttr = rowNode.getAttribute('r');
        const rowIndex = rowIndexAttr == null ? i : parseInt(rowIndexAttr, 10) - 1;
        if (!Number.isFinite(rowIndex) || rowIndex < 0) {
          continue;
        }
        while (rows.length <= rowIndex) {
          rows.push([]);
        }
        const rowValues = rows[rowIndex];
        const cellNodes = rowNode.getElementsByTagName('c');
        for (let j = 0; j < cellNodes.length; j++) {
          const cellNode = cellNodes[j];
          const ref = cellNode.getAttribute('r') || '';
          const columnIndex = columnLabelToIndex(ref);
          const value = readCellValue(cellNode, sharedStrings, stylesInfo);
          const targetIndex = columnIndex != null ? columnIndex : rowValues.length;
          rowValues[targetIndex] = value;
          if (targetIndex > maxColumnIndex) {
            maxColumnIndex = targetIndex;
          }
        }
      }

      const columnCount = maxColumnIndex + 1;
      if (columnCount === 0 || rows.length === 0) {
        return [];
      }

      const normalizedRows = rows.map(function (row) {
        const arr = new Array(columnCount);
        for (let col = 0; col < columnCount; col++) {
          if (!row || typeof row[col] === 'undefined') {
            arr[col] = '';
          } else {
            arr[col] = row[col];
          }
        }
        return arr;
      });

      const headerRow = normalizedRows[0] || [];
      const headers = headerRow.map(function (value) {
        if (value == null) {
          return '';
        }
        if (value instanceof Date) {
          return formatDateForApi(value);
        }
        return String(value).trim();
      });

      const result = [];
      for (let rowIndex = 1; rowIndex < normalizedRows.length; rowIndex++) {
        const row = normalizedRows[rowIndex];
        if (!row) {
          continue;
        }
        const entry = {};
        let hasData = false;
        for (let col = 0; col < headers.length; col++) {
          const header = headers[col];
          if (!header) {
            continue;
          }
          const cellValue = row[col];
          if (cellValue instanceof Date) {
            entry[header] = cellValue;
            hasData = true;
          } else if (cellValue == null || cellValue === '') {
            entry[header] = '';
          } else {
            entry[header] = cellValue;
            hasData = true;
          }
        }
        if (hasData) {
          result.push(entry);
        }
      }
      return result;
    }

    function createZipReader(arrayBuffer) {
      if (!(arrayBuffer instanceof ArrayBuffer)) {
        throw new Error('El archivo de Excel no es válido.');
      }
      if (typeof global.TextDecoder !== 'function') {
        throw new Error('Este navegador no soporta la lectura de archivos de Excel.');
      }
      if (typeof global.Uint8Array !== 'function') {
        throw new Error('Este navegador no soporta la lectura de archivos de Excel.');
      }
      const data = new Uint8Array(arrayBuffer);
      const textDecoder = new global.TextDecoder('utf-8');

      function readUint16(offset) {
        return data[offset] | (data[offset + 1] << 8);
      }

      function readUint32(offset) {
        return (
          data[offset] |
          (data[offset + 1] << 8) |
          (data[offset + 2] << 16) |
          (data[offset + 3] << 24)
        );
      }

      let eocdOffset = -1;
      const minOffset = Math.max(0, data.length - 0xFFFF);
      for (let i = data.length - 22; i >= minOffset; i--) {
        if (
          data[i] === 0x50 &&
          data[i + 1] === 0x4b &&
          data[i + 2] === 0x05 &&
          data[i + 3] === 0x06
        ) {
          eocdOffset = i;
          break;
        }
      }
      if (eocdOffset === -1) {
        throw new Error('El archivo de Excel está dañado o no se reconoce su formato.');
      }

      const entryCount = readUint16(eocdOffset + 10);
      const centralDirectoryOffset = readUint32(eocdOffset + 16);

      const entries = {};
      let offset = centralDirectoryOffset;
      for (let i = 0; i < entryCount; i++) {
        if (
          data[offset] !== 0x50 ||
          data[offset + 1] !== 0x4b ||
          data[offset + 2] !== 0x01 ||
          data[offset + 3] !== 0x02
        ) {
          throw new Error('El archivo de Excel está dañado (directorio central inválido).');
        }
        const compressionMethod = readUint16(offset + 10);
        const compressedSize = readUint32(offset + 20);
        const fileNameLength = readUint16(offset + 28);
        const extraLength = readUint16(offset + 30);
        const commentLength = readUint16(offset + 32);
        const localHeaderOffset = readUint32(offset + 42);
        const nameStart = offset + 46;
        const nameEnd = nameStart + fileNameLength;
        const rawName = data.slice(nameStart, nameEnd);
        const fileName = textDecoder.decode(rawName);
        entries[fileName] = {
          compressionMethod: compressionMethod,
          compressedSize: compressedSize,
          localHeaderOffset: localHeaderOffset
        };
        offset = nameEnd + extraLength + commentLength;
      }

      function getCompressedData(entry) {
        const localOffset = entry.localHeaderOffset;
        if (
          data[localOffset] !== 0x50 ||
          data[localOffset + 1] !== 0x4b ||
          data[localOffset + 2] !== 0x03 ||
          data[localOffset + 3] !== 0x04
        ) {
          throw new Error('El archivo de Excel está dañado (cabecera local inválida).');
        }
        const fileNameLength = readUint16(localOffset + 26);
        const extraFieldLength = readUint16(localOffset + 28);
        const dataStart = localOffset + 30 + fileNameLength + extraFieldLength;
        const dataEnd = dataStart + entry.compressedSize;
        return data.slice(dataStart, dataEnd);
      }

      async function decompressEntry(entry) {
        const compressed = getCompressedData(entry);
        if (entry.compressionMethod === 0) {
          return compressed.buffer.slice(compressed.byteOffset, compressed.byteOffset + compressed.byteLength);
        }
        if (entry.compressionMethod === 8) {
          if (typeof global.DecompressionStream !== 'function' || typeof global.Response !== 'function' || typeof global.Blob !== 'function') {
            throw new Error('Este navegador no soporta la descompresión de archivos ZIP requerida para leer Excel.');
          }
          const stream = new global.Response(
            new global.Blob([compressed])
          ).body.pipeThrough(new global.DecompressionStream('deflate-raw'));
          const buffer = await new global.Response(stream).arrayBuffer();
          return buffer;
        }
        throw new Error('El archivo de Excel usa un método de compresión no soportado.');
      }

      return {
        has: function (name) {
          return Object.prototype.hasOwnProperty.call(entries, name);
        },
        readText: async function (name) {
          if (!Object.prototype.hasOwnProperty.call(entries, name)) {
            throw new Error(`El archivo de Excel no contiene el recurso "${name}".`);
          }
          const buffer = await decompressEntry(entries[name]);
          return textDecoder.decode(new Uint8Array(buffer));
        }
      };
    }

    async function parseXlsxRows(arrayBuffer) {
      const zipReader = createZipReader(arrayBuffer);
      const workbookXml = await zipReader.readText('xl/workbook.xml').catch(function () {
        throw new Error('El archivo de Excel no contiene la información del libro.');
      });
      const workbookDoc = parseXmlDocument(workbookXml);
      const sheets = workbookDoc.getElementsByTagName('sheet');
      if (!sheets || sheets.length === 0) {
        throw new Error('El archivo de Excel no contiene hojas de cálculo.');
      }
      const firstSheet = sheets[0];
      const relId = firstSheet.getAttribute('r:id') || firstSheet.getAttribute('r:Id');
      if (!relId) {
        throw new Error('No se pudo determinar la hoja principal del archivo.');
      }
      const relsXml = await zipReader.readText('xl/_rels/workbook.xml.rels').catch(function () {
        throw new Error('El archivo de Excel no contiene la información de relaciones necesaria.');
      });
      const relsDoc = parseXmlDocument(relsXml);
      const relationshipNodes = relsDoc.getElementsByTagName('Relationship');
      let sheetTarget = '';
      for (let i = 0; i < relationshipNodes.length; i++) {
        const rel = relationshipNodes[i];
        const idAttr = rel.getAttribute('Id');
        if (idAttr === relId) {
          sheetTarget = rel.getAttribute('Target') || '';
          break;
        }
      }
      if (!sheetTarget) {
        throw new Error('No se encontró la hoja principal del archivo.');
      }
      const sheetPath = resolveSheetPath(sheetTarget);
      const sheetXml = await zipReader.readText(sheetPath).catch(function () {
        throw new Error('No se pudo leer la hoja principal del archivo.');
      });

      let sharedStrings = [];
      if (zipReader.has('xl/sharedStrings.xml')) {
        const sharedXml = await zipReader.readText('xl/sharedStrings.xml');
        sharedStrings = parseSharedStringsXml(sharedXml);
      }

      let stylesInfo = null;
      if (zipReader.has('xl/styles.xml')) {
        const stylesXml = await zipReader.readText('xl/styles.xml');
        stylesInfo = parseStylesXml(stylesXml);
      }

      return extractSheetRows(sheetXml, sharedStrings, stylesInfo);
    }

    function readExcelFile(file) {
      return new Promise(function (resolve, reject) {
        if (!file) {
          reject(new Error('Selecciona un archivo de Excel para continuar.'));
          return;
        }
        const reader = new global.FileReader();
        reader.onerror = function () {
          reject(new Error('No se pudo leer el archivo seleccionado.'));
        };
        reader.onload = function (event) {
          try {
            const result = event && event.target ? event.target.result : null;
            if (!(result instanceof ArrayBuffer)) {
              throw new Error('El archivo de Excel no se pudo interpretar.');
            }
            parseXlsxRows(result).then(resolve).catch(function (err) {
              reject(err instanceof Error ? err : new Error('No se pudo procesar el archivo de Excel.'));
            });
          } catch (err) {
            reject(err instanceof Error ? err : new Error('No se pudo procesar el archivo de Excel.'));
          }
        };
        reader.readAsArrayBuffer(file);
      });
    }

    function prepareBulkRows(rawRows) {
      const issues = [];
      const prepared = [];
      if (!Array.isArray(rawRows) || rawRows.length === 0) {
        return { rows: [], issues: ['El archivo no contiene filas con datos.'] };
      }

      const headerSet = new Set();
      rawRows.forEach(function (row) {
        if (!row || typeof row !== 'object') {
          return;
        }
        Object.keys(row).forEach(function (key) {
          const canonical = normalizeBulkHeader(key);
          if (canonical) {
            headerSet.add(canonical);
          }
        });
      });

      const missingHeaders = BULK_REQUIRED_HEADERS.filter(function (label) {
        return !headerSet.has(label);
      });
      if (missingHeaders.length > 0) {
        return {
          rows: [],
          issues: [`Faltan las columnas obligatorias: ${missingHeaders.join(', ')}.`]
        };
      }

      rawRows.forEach(function (row, index) {
        if (!row || typeof row !== 'object') {
          return;
        }

        const normalizedValues = {};
        let hasAnyValue = false;
        Object.keys(row).forEach(function (key) {
          const canonical = normalizeBulkHeader(key);
          if (!canonical) {
            return;
          }
          const value = row[key];
          if (value instanceof Date && !isNaN(value.getTime())) {
            hasAnyValue = true;
          } else if (value != null && String(value).trim() !== '') {
            hasAnyValue = true;
          }
          normalizedValues[canonical] = value;
        });

        if (!hasAnyValue) {
          return;
        }

        const rowNumber = index + 2;
        const rowIssues = [];
        const output = {};

        const rawTrip = normalizedValues.Trip;
        const tripValue = rawTrip == null ? '' : String(rawTrip).trim();
        if (!tripValue) {
          rowIssues.push('Trip vacío');
        } else if (!/^\d+$/.test(tripValue)) {
          rowIssues.push('Trip inválido');
        } else if (Number(tripValue) < 225000) {
          rowIssues.push('Trip menor a 225000');
        }
        output.Trip = tripValue;

        const rawEjecutivo = normalizedValues.Ejecutivo;
        const ejecutivoValue = rawEjecutivo == null ? '' : String(rawEjecutivo).trim();
        if (!ejecutivoValue) {
          rowIssues.push('Ejecutivo vacío');
        }
        output.Ejecutivo = ejecutivoValue;

        ['Caja', 'Estatus', 'Cliente', 'Destino', 'Referencia', 'Segmento'].forEach(function (label) {
          const rawValue = normalizedValues[label];
          const value = rawValue == null ? '' : String(rawValue).trim();
          output[label] = value;
        });

        BULK_DATE_HEADER_SET.forEach(function (label) {
          const conversion = convertBulkDateValue(normalizedValues[label]);
          if (conversion.error) {
            rowIssues.push(`${label}: ${conversion.error}`);
          }
          output[label] = conversion.value;
        });

        if (rowIssues.length === 0) {
          prepared.push(output);
        } else {
          issues.push(`Fila ${rowNumber}: ${rowIssues.join(', ')}.`);
        }
      });

      return { rows: prepared, issues: issues };
    }

    function setStatus(message, type) {
      const el = refs.status;
      if (!el) return;
      if (!message) {
        el.textContent = '';
        el.className = 'sheet-status';
        el.hidden = true;
        el.removeAttribute('title');
        return;
      }
      const statusClass = 'sheet-status' + (type ? ` is-${type}` : '');
      el.className = statusClass;
      el.hidden = false;
      el.textContent = message;
      el.title = message;
    }

    async function processBulkUploadFile(file) {
      if (!file) {
        setStatus('Selecciona un archivo de Excel para continuar.', 'error');
        return;
      }
      const extension = getFileExtension(file.name);
      if (!BULK_ALLOWED_EXTENSIONS.has(extension)) {
        setStatus('El archivo debe estar en formato .xlsx o .xlsm.', 'error');
        return;
      }
      if (!state.currentUser) {
        setStatus('Inicia sesión para cargar registros.', 'error');
        showLoginModal();
        return;
      }
      if (!state.token) {
        setStatus('Sesión expirada. Inicia sesión nuevamente.', 'error');
        showLoginModal();
        return;
      }
      if (!state.config || !state.config.API_BASE) {
        setStatus('Falta configurar la URL del Apps Script.', 'error');
        return;
      }
      if (state.loading) {
        setStatus('Espera a que finalice la operación en curso.', 'info');
        return;
      }

      toggleLoading(true);
      try {
        setStatus('Procesando archivo de Excel…', 'info');
        const rawRows = await readExcelFile(file);
        const preparation = prepareBulkRows(rawRows);
        const validRows = Array.isArray(preparation.rows) ? preparation.rows : [];
        const issues = Array.isArray(preparation.issues) ? preparation.issues : [];

        if (validRows.length === 0) {
          const issueSummary = summarizeIssues(issues, 3);
          const message = issueSummary
            ? `No se encontraron filas válidas. ${issueSummary}`
            : 'No se encontraron filas válidas en el archivo.';
          setStatus(message, 'error');
          return;
        }

        const response = await submitBulkAddRequest(state.config.API_BASE, state.token, validRows);
        if (response && response.success === false) {
          throw new Error(response.error || 'No se pudo completar la carga masiva.');
        }
        const insertedCount = typeof response.inserted === 'number' ? response.inserted : validRows.length;
        const duplicates = Array.isArray(response.duplicates) ? response.duplicates : [];

        if (insertedCount > 0) {
          await loadData();
        }

        const parts = [];
        const insertedLabel = insertedCount === 1 ? 'registro' : 'registros';
        parts.push(`Se agregaron ${insertedCount} ${insertedLabel}.`);

        if (duplicates.length > 0) {
          const duplicatesLabel = duplicates.length === 1 ? 'trip' : 'trips';
          const duplicatesSummary = summarizeValues(duplicates, 5);
          if (duplicatesSummary) {
            parts.push(`Se omitieron ${duplicates.length} ${duplicatesLabel} por duplicado: ${duplicatesSummary}.`);
          } else {
            parts.push(`Se omitieron ${duplicates.length} ${duplicatesLabel} por duplicado.`);
          }
        }

        if (issues.length > 0) {
          const issuesLabel = issues.length === 1 ? 'fila' : 'filas';
          const issuesSummary = summarizeIssues(issues, 3);
          if (issuesSummary) {
            const trimmedSummary = issuesSummary.replace(/[.]+$/g, '');
            parts.push(`Se ignoraron ${issues.length} ${issuesLabel} por errores: ${trimmedSummary}.`);
          } else {
            parts.push(`Se ignoraron ${issues.length} ${issuesLabel} por errores.`);
          }
        }

        setStatus(parts.join(' '), 'success');
      } catch (err) {
        const fallbackMessage = 'Error al procesar el archivo de Excel.';
        const message = err && err.message ? err.message : fallbackMessage;
        setStatus(message, 'error');
        if (err && err.status === 401) {
          state.token = '';
          setStoredValue(STORAGE_TOKEN_KEY, null);
          showLoginModal();
        }
      } finally {
        toggleLoading(false);
      }
    }

    function handleBulkUploadClick() {
      if (!refs.bulkUploadInput) {
        setStatus('No se pudo iniciar la carga masiva.', 'error');
        return;
      }
      if (!state.currentUser) {
        setStatus('Inicia sesión para cargar registros.', 'error');
        showLoginModal();
        return;
      }
      if (state.loading) {
        setStatus('Espera a que finalice la operación en curso.', 'info');
        return;
      }
      refs.bulkUploadInput.value = '';
      refs.bulkUploadInput.click();
    }

    async function handleBulkUploadInputChange(event) {
      const input = event && event.target ? event.target : null;
      const files = input && input.files ? input.files : null;
      const file = files && files[0] ? files[0] : null;
      try {
        if (file) {
          await processBulkUploadFile(file);
        }
      } finally {
        if (input) {
          input.value = '';
        }
      }
    }

    function setEditModalMode(mode) {
      const content = EDIT_MODAL_CONTENT[mode] || EDIT_MODAL_CONTENT.edit;
      if (refs.editTitle) {
        refs.editTitle.textContent = content.title;
      }
      if (refs.editHint) {
        refs.editHint.textContent = content.hint;
      }
      if (refs.editSubmitButton) {
        refs.editSubmitButton.textContent = content.submit;
      }
    }

    function populateEditFormValues(values) {
      if (!refs.editForm || !values) {
        return;
      }
      Object.keys(values).forEach(function (key) {
        const input = refs.editForm.querySelector('[name="' + key + '"]');
        if (!input) {
          return;
        }
        const rawValue = values[key];
        let preparedValue;
        if (DATE_FIELD_SET.has(key)) {
          preparedValue = toDateInputValue(rawValue);
        } else if (rawValue == null) {
          preparedValue = '';
        } else {
          preparedValue = String(rawValue);
        }
        if (input.tagName === 'INPUT' || input.tagName === 'TEXTAREA') {
          input.value = preparedValue || '';
        }
      });
    }

    function syncDateInputsWithState(range) {
      const normalized = normalizeDateRange(range || {});
      if (refs.dateStartInput) {
        refs.dateStartInput.value = dateToInputDateValue(normalized.start);
      }
      if (refs.dateEndInput) {
        refs.dateEndInput.value = dateToInputDateValue(normalized.end);
      }
    }

    function normalizeStatusValue(value) {
      if (value == null) {
        return '';
      }
      return String(value).trim();
    }

    function getStatusBadgeSlug(value) {
      const normalized = normalizeStatusValue(value);
      if (!normalized) {
        return '';
      }
      return normalized
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, '-')
        .replace(/^-+|-+$/g, '');
    }

    function isSameStatus(a, b) {
      return normalizeStatusValue(a).toLowerCase() === normalizeStatusValue(b).toLowerCase();
    }

    function createStatusOptionElement(value, label, isSelected) {
      const item = doc.createElement('li');
      item.className = 'status-filter__option';
      const button = doc.createElement('button');
      button.type = 'button';
      button.className =
        'status-filter__option-button' + (isSelected ? ' is-selected' : '');
      button.setAttribute('data-status-value', value);
      button.setAttribute('role', 'option');
      button.setAttribute('aria-selected', isSelected ? 'true' : 'false');
      button.textContent = label == null || label === '' ? '—' : label;
      item.appendChild(button);
      return item;
    }

    function updateAvailableStatuses(values) {
      const locale = state.locale || DEFAULT_LOCALE;
      const unique = new Map();
      if (Array.isArray(values)) {
        values.forEach(function (value) {
          const normalized = normalizeStatusValue(value);
          if (!normalized) {
            return;
          }
          const key = normalized.toLowerCase();
          if (!unique.has(key)) {
            unique.set(key, normalized);
          }
        });
      }
      const sorted = Array.from(unique.values()).sort(function (a, b) {
        return a.localeCompare(b, locale, { sensitivity: 'base' });
      });
      const previous = state.availableStatuses || [];
      const changed =
        sorted.length !== previous.length ||
        sorted.some(function (value, index) {
          return value !== previous[index];
        });
      if (changed) {
        state.availableStatuses = sorted;
      }

      const selected = normalizeStatusValue(state.filters.status);
      let canonicalSelected = '';
      if (selected) {
        const match = sorted.find(function (value) {
          return isSameStatus(value, selected);
        });
        if (match) {
          canonicalSelected = match;
        }
      }

      if (selected && !canonicalSelected) {
        state.filters.status = '';
      } else if (canonicalSelected && state.filters.status !== canonicalSelected) {
        state.filters.status = canonicalSelected;
      }

      if (state.availableStatuses.length === 0 && state.isStatusPopoverOpen) {
        state.isStatusPopoverOpen = false;
      }

      renderStatusFilter();
    }

    function renderStatusFilter() {
      const selected = normalizeStatusValue(state.filters.status);
      if (refs.statusLabel) {
        refs.statusLabel.textContent = selected ? state.filters.status : 'Todos';
      }

      if (refs.statusFilter) {
        if (state.isStatusPopoverOpen && state.availableStatuses.length > 0) {
          refs.statusFilter.classList.add('is-open');
        } else {
          refs.statusFilter.classList.remove('is-open');
        }
        refs.statusFilter.setAttribute('data-has-selection', selected ? 'true' : 'false');
      }

      if (refs.statusToggleButton) {
        refs.statusToggleButton.setAttribute('aria-expanded', state.isStatusPopoverOpen ? 'true' : 'false');
        refs.statusToggleButton.disabled = state.availableStatuses.length === 0;
      }

      if (refs.statusPopover) {
        const shouldShow = state.isStatusPopoverOpen && state.availableStatuses.length > 0;
        refs.statusPopover.hidden = !shouldShow;
      }

      if (refs.statusList) {
        const list = refs.statusList;
        list.innerHTML = '';
        if (state.availableStatuses.length === 0) {
          const emptyItem = doc.createElement('li');
          emptyItem.className = 'status-filter__empty';
          emptyItem.setAttribute('role', 'presentation');
          emptyItem.textContent = 'Sin estatus disponibles';
          list.appendChild(emptyItem);
        } else {
          const normalizedSelected = selected;
          list.appendChild(createStatusOptionElement('', 'Todos', normalizedSelected === ''));
          state.availableStatuses.forEach(function (status) {
            const isSelected = normalizedSelected !== '' && isSameStatus(status, normalizedSelected);
            list.appendChild(createStatusOptionElement(status, status, isSelected));
          });
        }
      }

      if (state.isStatusPopoverOpen && !wasStatusPopoverOpen && refs.statusList) {
        const focusTarget =
          refs.statusList.querySelector('.status-filter__option-button.is-selected') ||
          refs.statusList.querySelector('.status-filter__option-button');
        if (focusTarget && typeof focusTarget.focus === 'function') {
          focusTarget.focus();
        }
      }

      wasStatusPopoverOpen = state.isStatusPopoverOpen;
    }

    function openStatusPopover() {
      if (state.isStatusPopoverOpen || state.availableStatuses.length === 0) {
        return;
      }
      state.isStatusPopoverOpen = true;
      renderStatusFilter();
    }

    function closeStatusPopover() {
      if (!state.isStatusPopoverOpen) {
        return;
      }
      state.isStatusPopoverOpen = false;
      renderStatusFilter();
    }

    function toggleStatusPopover() {
      if (state.isStatusPopoverOpen) {
        closeStatusPopover();
      } else {
        openStatusPopover();
      }
    }

    function applyStatusFilter(value) {
      const normalized = normalizeStatusValue(value);
      let nextValue = '';
      if (normalized) {
        const match = (state.availableStatuses || []).find(function (status) {
          return isSameStatus(status, normalized);
        });
        nextValue = match || normalized;
      }
      if (state.filters.status === nextValue) {
        return;
      }
      state.filters.status = nextValue;
      renderTable();
    }


    function isDateFilterEnabledForCurrentView() {
      const activeView = getActiveView();
      return Boolean(activeView && activeView.dateFilterEnabled);
    }

    function renderDateFilter() {
      const normalized = normalizeDateRange(state.filters.dateRange || {});
      state.filters.dateRange = normalized;
      const hasRange = isValidDate(normalized.start) || isValidDate(normalized.end);

      if (refs.dateLabel) {
        refs.dateLabel.textContent = formatDateRangeLabel(normalized, state.locale);
      }

      const shouldShowDateFilter = isDateFilterEnabledForCurrentView();
      const filterCard = refs.dateFilter && typeof refs.dateFilter.closest === 'function'
        ? refs.dateFilter.closest('.filter-card')
        : null;

      if (filterCard) {
        if (shouldShowDateFilter) {
          filterCard.classList.remove('is-hidden');
          filterCard.removeAttribute('hidden');
        } else {
          filterCard.classList.add('is-hidden');
          filterCard.setAttribute('hidden', '');
        }
      } else if (refs.dateFilter) {
        if (shouldShowDateFilter) {
          refs.dateFilter.classList.remove('is-hidden');
          refs.dateFilter.removeAttribute('hidden');
        } else {
          refs.dateFilter.classList.add('is-hidden');
          refs.dateFilter.setAttribute('hidden', '');
        }
      }

      if (!shouldShowDateFilter) {
        state.isDatePopoverOpen = false;
        if (refs.dateFilter) {
          refs.dateFilter.classList.remove('is-open');
          refs.dateFilter.setAttribute('data-has-range', hasRange ? 'true' : 'false');
        }
        if (refs.datePopover) {
          refs.datePopover.hidden = true;
        }
        if (refs.dateToggleButton) {
          refs.dateToggleButton.setAttribute('aria-expanded', 'false');
        }
        wasDatePopoverOpen = false;
        return;
      }

      if (refs.dateFilter) {
        if (state.isDatePopoverOpen) {
          refs.dateFilter.classList.add('is-open');
        } else {
          refs.dateFilter.classList.remove('is-open');
        }
        refs.dateFilter.setAttribute('data-has-range', hasRange ? 'true' : 'false');
      }

      if (refs.datePopover) {
        refs.datePopover.hidden = !state.isDatePopoverOpen;
      }

      if (refs.dateToggleButton) {
        refs.dateToggleButton.setAttribute('aria-expanded', state.isDatePopoverOpen ? 'true' : 'false');
      }

      syncDateInputsWithState(normalized);

      if (state.isDatePopoverOpen && !wasDatePopoverOpen && refs.dateStartInput) {
        refs.dateStartInput.focus();
      }

      wasDatePopoverOpen = state.isDatePopoverOpen;
    }

    function openDatePopover() {
      if (state.isDatePopoverOpen) {
        return;
      }
      state.isDatePopoverOpen = true;
      renderDateFilter();
    }

    function closeDatePopover() {
      if (!state.isDatePopoverOpen) {
        return;
      }
      state.isDatePopoverOpen = false;
      renderDateFilter();
    }

    function toggleDatePopover() {
      if (state.isDatePopoverOpen) {
        closeDatePopover();
      } else {
        openDatePopover();
      }
    }

    function applyDateRange(range) {
      const currentRange = normalizeDateRange(state.filters.dateRange || {});
      const nextRange = normalizeDateRange(range || {});
      if (areRangesEqual(currentRange, nextRange)) {
        state.filters.dateRange = currentRange;
        renderDateFilter();
        return;
      }
      state.filters.dateRange = nextRange;
      renderDateFilter();
      renderTable();
    }

    function shiftCurrentDateRange(days) {
      const normalized = normalizeDateRange(state.filters.dateRange || {});
      const hasRange = isValidDate(normalized.start) || isValidDate(normalized.end);
      const baseRange = hasRange ? normalized : createTodayRange();
      const delta = Number(days) * MS_PER_DAY;
      const nextStart = isValidDate(baseRange.start) ? new Date(baseRange.start.getTime() + delta) : null;
      const nextEnd = isValidDate(baseRange.end) ? new Date(baseRange.end.getTime() + delta) : null;
      applyDateRange({ start: nextStart, end: nextEnd });
    }

    function handleDateInputChange() {
      const startValue = refs.dateStartInput ? refs.dateStartInput.value : '';
      const endValue = refs.dateEndInput ? refs.dateEndInput.value : '';
      const startDate = parseInputDateValue(startValue);
      const endDate = parseInputDateValue(endValue);
      applyDateRange({ start: startDate, end: endDate });
    }

    function handleDatePrev(event) {
      if (event) {
        event.preventDefault();
      }
      shiftCurrentDateRange(-1);
    }

    function handleDateNext(event) {
      if (event) {
        event.preventDefault();
      }
      shiftCurrentDateRange(1);
    }

    function handleDateToggle(event) {
      if (event) {
        event.preventDefault();
      }
      toggleDatePopover();
    }

    function handleDateClear(event) {
      if (event) {
        event.preventDefault();
      }
      applyDateRange({ start: null, end: null });
      closeDatePopover();
    }

    function handleStatusToggle(event) {
      if (event) {
        event.preventDefault();
      }
      toggleStatusPopover();
    }

    function handleStatusListClick(event) {
      const target = event && event.target ? event.target : null;
      if (!target || typeof target.closest !== 'function') {
        return;
      }
      const button = target.closest('[data-status-value]');
      if (!button) {
        return;
      }
      event.preventDefault();
      const value = button.getAttribute('data-status-value') || '';
      applyStatusFilter(value);
      closeStatusPopover();
    }

    function handleDocumentClick(event) {
      const target = event && event.target ? event.target : null;
      if (state.isDatePopoverOpen) {
        const dateContainer = refs.dateFilter;
        if (!dateContainer || !(target && dateContainer.contains(target))) {
          closeDatePopover();
        }
      }
      if (state.isStatusPopoverOpen) {
        const statusContainer = refs.statusFilter;
        if (!statusContainer || !(target && statusContainer.contains(target))) {
          closeStatusPopover();
        }
      }
    }

    function handleDocumentKeydown(event) {
      if (!state.isDatePopoverOpen && !state.isStatusPopoverOpen) {
        return;
      }
      if (event && (event.key === 'Escape' || event.key === 'Esc')) {
        if (state.isDatePopoverOpen) {
          closeDatePopover();
        }
        if (state.isStatusPopoverOpen) {
          closeStatusPopover();
        }
      }
    }

    function resetFilters() {
      state.filters.searchText = '';
      state.filters.dateRange = { start: null, end: null };
      state.filters.status = '';
      state.isDatePopoverOpen = false;
      state.isStatusPopoverOpen = false;
      state.availableStatuses = [];
      if (refs.filterSearchInput) {
        refs.filterSearchInput.value = '';
      }
      renderDateFilter();
      renderStatusFilter();
    }

    function getActiveView() {
      const currentId = state.currentViewId;
      for (let i = 0; i < TABLE_VIEWS.length; i++) {
        const view = TABLE_VIEWS[i];
        if (view && view.id === currentId) {
          return view;
        }
      }
      return TABLE_VIEWS[0];
    }

    function renderViewMenu() {
      const container = refs.viewMenu;
      if (!container) {
        return;
      }
      container.innerHTML = '';
      const fragment = doc.createDocumentFragment();
      const label = doc.createElement('span');
      label.className = 'sheet-grid__views-label';
      label.textContent = 'Vistas:';
      fragment.appendChild(label);
      TABLE_VIEWS.forEach(function (view) {
        if (!view) {
          return;
        }
        const button = doc.createElement('button');
        button.type = 'button';
        button.className =
          'sheet-grid__view-button' + (view.id === state.currentViewId ? ' is-active' : '');
        button.setAttribute('data-view-id', view.id);
        button.setAttribute('aria-pressed', view.id === state.currentViewId ? 'true' : 'false');
        button.textContent = view.label;
        fragment.appendChild(button);
      });
      container.appendChild(fragment);
    }

    function setCurrentView(viewId) {
      const exists = TABLE_VIEWS.some(function (view) {
        return view && view.id === viewId;
      });
      const nextId = exists ? viewId : state.currentViewId;
      if (!nextId) {
        return;
      }
      if (state.currentViewId === nextId) {
        renderViewMenu();
        return;
      }
      state.currentViewId = nextId;
      resetFilters();
      renderViewMenu();
      renderTable();
      renderDateFilter();
    }

    function handleViewMenuClick(event) {
      const target = event.target;
      if (!target || typeof target.closest !== 'function') {
        return;
      }
      const button = target.closest('[data-view-id]');
      if (!button) {
        return;
      }
      event.preventDefault();
      const viewId = button.getAttribute('data-view-id');
      if (viewId) {
        setCurrentView(viewId);
      }
    }

    function handleFilterSearchInput(event) {
      const target = event && event.target ? event.target : null;
      const value = target && target.value != null ? String(target.value) : '';
      if (state.filters.searchText === value) {
        return;
      }
      state.filters.searchText = value;
      renderTable();
    }

    function showBackdrop() {
      if (!refs.backdrop) {
        return;
      }
      refs.backdrop.classList.remove('hidden');
      refs.backdrop.classList.add('is-visible');
    }

    function hideBackdropIfNoModalVisible() {
      if (!refs.backdrop) {
        return;
      }
      const loginVisible = refs.loginModal && refs.loginModal.classList.contains('is-visible');
      const editVisible = refs.editModal && refs.editModal.classList.contains('is-visible');
      if (!loginVisible && !editVisible) {
        refs.backdrop.classList.remove('is-visible');
        refs.backdrop.classList.add('hidden');
      }
    }

    function setEditFormDisabled(isDisabled) {
      if (!refs.editForm) {
        return;
      }
      const elements = refs.editForm.querySelectorAll('input, textarea, button');
      elements.forEach(function (element) {
        element.disabled = Boolean(isDisabled);
      });
    }

    function updateLastUpdated(date) {
      if (!refs.lastUpdated) return;
      if (!date) {
        refs.lastUpdated.textContent = 'Última actualización: —';
        return;
      }
      const formatted = fmtDate(date, state.locale);
      refs.lastUpdated.textContent = `Última actualización: ${formatted}`;
    }

    function updateUserBadge() {
      if (!refs.currentUser) return;
      if (!state.currentUser) {
        refs.currentUser.textContent = '';
        return;
      }
      refs.currentUser.textContent = `Usuario: ${state.currentUser.displayName || state.currentUser.username}`;
    }

    function toggleLoading(isLoading) {
      state.loading = isLoading;
      if (refs.refreshButton) {
        refs.refreshButton.disabled = Boolean(isLoading);
      }
      if (isLoading) {
        appRoot.classList.add('is-loading');
      } else {
        appRoot.classList.remove('is-loading');
      }
    }

    function clearTable() {
      if (refs.tableHead) refs.tableHead.innerHTML = '';
      if (refs.tableBody) refs.tableBody.innerHTML = '';
    }

    function renderTable() {
      if (!refs.tableHead || !refs.tableBody) {
        return;
      }

      clearTable();

      if (!Array.isArray(state.data) || state.data.length === 0) {
        updateAvailableStatuses([]);
        setStatus('No hay datos disponibles en la hoja.', 'info');
        return;
      }

      const headers = state.data[0] || [];
      const columnKeys = headers.map(function (header) {
        return getColumnKeyFromHeader(header);
      });
      const dataRows = state.data.slice(1);
      let columnCount = headers.length;
      for (let i = 0; i < dataRows.length; i++) {
        if (Array.isArray(dataRows[i]) && dataRows[i].length > columnCount) {
          columnCount = dataRows[i].length;
        }
      }

      const columnMap = {};
      const dateColumnIndices = [];
      headers.forEach(function (header, index) {
        const key = columnKeys[index];
        if (key) {
          columnMap[key] = index;
        }
        if (isDateHeader(header)) {
          dateColumnIndices.push(index);
        }
      });

      const statusColumnIndex =
        typeof columnMap.estatus === 'number' && columnMap.estatus >= 0 ? columnMap.estatus : null;
      if (statusColumnIndex != null) {
        const statuses = dataRows.reduce(function (acc, row) {
          if (Array.isArray(row) && statusColumnIndex < row.length) {
            acc.push(row[statusColumnIndex]);
          }
          return acc;
        }, []);
        updateAvailableStatuses(statuses);
      } else {
        updateAvailableStatuses([]);
      }

      const llegadaCargaIndex =
        typeof columnMap.llegadaCarga === 'number' && columnMap.llegadaCarga >= 0 ? columnMap.llegadaCarga : null;

      const rowsWithIndex = dataRows.map(function (row, index) {
        return {
          row: row,
          dataIndex: index + 1
        };
      });

      const activeView = getActiveView();
      const filterContext = {
        columnMap: columnMap,
        headers: headers,
        now: new Date()
      };
      const referenceNow = isValidDate(filterContext.now) ? filterContext.now : new Date();
      let rowsToRender = rowsWithIndex;
      if (activeView && typeof activeView.filter === 'function') {
        rowsToRender = rowsWithIndex.filter(function (entry) {
          try {
            return activeView.filter(entry.row, filterContext);
          } catch (err) {
            return false;
          }
        });
      }

      const normalizedDateRange = normalizeDateRange(state.filters.dateRange || {});
      const rangeChanged = !areRangesEqual(state.filters.dateRange || {}, normalizedDateRange);
      state.filters.dateRange = normalizedDateRange;
      if (rangeChanged) {
        renderDateFilter();
      }

      const startTime = isValidDate(normalizedDateRange.start) ? normalizedDateRange.start.getTime() : null;
      const endTime = isValidDate(normalizedDateRange.end) ? normalizedDateRange.end.getTime() : null;
      const citaCargaIndex = typeof columnMap.citaCarga === 'number' && columnMap.citaCarga >= 0 ? columnMap.citaCarga : null;
      const dateFilterIndices = citaCargaIndex != null ? [citaCargaIndex] : [];
      const shouldApplyDateFilter = isDateFilterEnabledForCurrentView();
      const hasDateRange = shouldApplyDateFilter && (startTime != null || endTime != null) && dateFilterIndices.length > 0;

      if (hasDateRange) {
        rowsToRender = rowsToRender.filter(function (entry) {
          const row = Array.isArray(entry.row) ? entry.row : [];
          for (let i = 0; i < dateFilterIndices.length; i++) {
            const columnIndex = dateFilterIndices[i];
            if (columnIndex >= row.length) {
              continue;
            }
            const cellValue = row[columnIndex];
            if (cellValue == null || cellValue === '') {
              continue;
            }
            const cellDate = parseDateValue(cellValue);
            if (!cellDate) {
              continue;
            }
            const cellTime = cellDate.getTime();
            if ((startTime == null || cellTime >= startTime) && (endTime == null || cellTime <= endTime)) {
              return true;
            }
          }
          return false;
        });
      }

      const selectedStatus = normalizeStatusValue(state.filters.status);
      if (statusColumnIndex != null && selectedStatus) {
        rowsToRender = rowsToRender.filter(function (entry) {
          const row = Array.isArray(entry.row) ? entry.row : [];
          if (statusColumnIndex >= row.length) {
            return false;
          }
          const cellValue = row[statusColumnIndex];
          if (cellValue == null) {
            return false;
          }
          return isSameStatus(cellValue, selectedStatus);
        });
      }


      const searchQuery = String(state.filters.searchText || '').trim().toLowerCase();
      if (searchQuery) {
        const searchableKeys = ['trip', 'caja', 'referencia', 'cliente'];
        const searchableIndices = searchableKeys
          .map(function (key) {
            return columnMap[key];
          })
          .filter(function (index) {
            return typeof index === 'number' && index >= 0;
          });
        if (searchableIndices.length > 0) {
          rowsToRender = rowsToRender.filter(function (entry) {
            const row = Array.isArray(entry.row) ? entry.row : [];
            return searchableIndices.some(function (index) {
              if (index >= row.length) {
                return false;
              }
              const cellValue = row[index];
              if (cellValue == null) {
                return false;
              }
              return String(cellValue).toLowerCase().indexOf(searchQuery) !== -1;
            });
          });
        }
      }

      const sortColumnIndices = (function () {
        if (dateColumnIndices.length === 0) {
          return [];
        }
        const indices = dateColumnIndices.slice();
        if (activeView && activeView.id === 'daily-loads') {
          const citaCargaIndex = columnMap.citaCarga;
          if (typeof citaCargaIndex === 'number' && citaCargaIndex >= 0) {
            const filtered = indices.filter(function (value) {
              return value !== citaCargaIndex;
            });
            filtered.unshift(citaCargaIndex);
            return filtered;
          }
        }
        return indices;
      })();

      if (sortColumnIndices.length > 0 && rowsToRender.length > 1) {
        const sortableEntries = rowsToRender.map(function (entry) {
          const row = Array.isArray(entry.row) ? entry.row : [];
          const sortValues = sortColumnIndices.map(function (columnIndex) {
            if (columnIndex >= row.length) {
              return Number.POSITIVE_INFINITY;
            }
            return getDateSortValue(row[columnIndex]);
          });
          return {
            entry: entry,
            sortValues: sortValues
          };
        });

        sortableEntries.sort(function (a, b) {
          for (let i = 0; i < sortColumnIndices.length; i++) {
            const aValue = a.sortValues[i];
            const bValue = b.sortValues[i];
            if (aValue < bValue) {
              return -1;
            }
            if (aValue > bValue) {
              return 1;
            }
          }
          return a.entry.dataIndex - b.entry.dataIndex;
        });

        rowsToRender = sortableEntries.map(function (item) {
          return item.entry;
        });
      }

      const headerRow = doc.createElement('tr');

      for (let c = 0; c < columnCount; c++) {
        const th = doc.createElement('th');
        const headerLabel = headers[c];
        const label = headerLabel != null && headerLabel !== '' ? headerLabel : columnLetter(c);
        th.textContent = formatHeaderLabel(label);
        if (isDateHeader(label)) {
          th.classList.add('is-date');
        }
        const columnKey = c < columnKeys.length ? columnKeys[c] : null;
        if (columnKey && NOWRAP_COLUMN_KEYS.has(columnKey)) {
          th.classList.add('is-nowrap');
        }
        headerRow.appendChild(th);
      }

      refs.tableHead.appendChild(headerRow);

      const fragment = doc.createDocumentFragment();
      rowsToRender.forEach(function (entry) {
        const row = Array.isArray(entry.row) ? entry.row : [];
        const tr = doc.createElement('tr');
        let shouldHighlightCitaCarga = false;

        if (citaCargaIndex != null && citaCargaIndex < row.length) {
          const rawCitaCarga = row[citaCargaIndex];
          if (rawCitaCarga != null && rawCitaCarga !== '') {
            const citaCargaDate = parseDateValue(rawCitaCarga);
            if (isValidDate(citaCargaDate) && citaCargaDate.getTime() < referenceNow.getTime()) {
              const rawLlegadaCarga =
                llegadaCargaIndex != null && llegadaCargaIndex < row.length
                  ? row[llegadaCargaIndex]
                  : null;
              const hasLlegadaCarga = rawLlegadaCarga != null && String(rawLlegadaCarga).trim() !== '';
              if (!hasLlegadaCarga) {
                shouldHighlightCitaCarga = true;
              }
            }
          }
        }

        if (shouldHighlightCitaCarga) {
          tr.classList.add('is-cita-carga-vencida');
        }

        for (let c = 0; c < columnCount; c++) {
          const td = doc.createElement('td');
          const headerLabel = headers[c];
          const columnKey = c < columnKeys.length ? columnKeys[c] : null;
          let value = row && row[c] != null ? row[c] : '';
          const isTripColumn = columnKey === 'trip';
          const isTrackingColumn = columnKey === 'tracking';
          const isStatusColumn = columnKey === 'estatus';
          if (isDateHeader(headerLabel) && value !== '') {
            const formatted = fmtDate(value, state.locale);
            value = formatted || value;
            td.classList.add('is-date');
          }
          if (columnKey && NOWRAP_COLUMN_KEYS.has(columnKey)) {
            td.classList.add('is-nowrap');
          }
          if (value === null || value === undefined || value === '') {
            td.classList.add('is-empty');
            td.textContent = '';
          } else if (isTripColumn) {
            const displayValue = typeof value === 'string' ? value : String(value);
            if (displayValue) {
              const button = doc.createElement('button');
              button.type = 'button';
              button.className = 'table-link-button';
              button.setAttribute('data-action', 'open-edit');
              button.setAttribute('data-row-index', String(entry.dataIndex));
              button.setAttribute('aria-label', `Editar registro del trip ${displayValue}`);
              button.title = 'Editar registro';
              button.textContent = displayValue;
              td.appendChild(button);
            } else {
              td.classList.add('is-empty');
              td.textContent = '';
            }
          } else if (isTrackingColumn) {
            const displayValue = typeof value === 'string' ? value.trim() : String(value).trim();
            if (displayValue) {
              const link = doc.createElement('a');
              link.href = displayValue;
              link.target = '_blank';
              link.rel = 'noopener noreferrer';
              link.className = 'table-link table-link--icon';
              link.setAttribute('aria-label', 'Abrir tracking en una nueva pestaña');
              link.title = 'Abrir tracking';

              const svgNamespace = 'http://www.w3.org/2000/svg';
              const icon = doc.createElementNS(svgNamespace, 'svg');
              icon.classList.add('table-link__icon');
              icon.setAttribute('aria-hidden', 'true');
              icon.setAttribute('focusable', 'false');
              icon.setAttribute('viewBox', '0 0 24 24');
              icon.setAttribute('width', '24');
              icon.setAttribute('height', '24');

              const createPath = (d) => {
                const path = doc.createElementNS(svgNamespace, 'path');
                path.setAttribute('d', d);
                path.setAttribute('fill', 'none');
                path.setAttribute('stroke', 'currentColor');
                path.setAttribute('stroke-width', '1.8');
                path.setAttribute('stroke-linecap', 'round');
                path.setAttribute('stroke-linejoin', 'round');
                return path;
              };

              icon.appendChild(
                createPath('M10.59 13.41a2 2 0 0 1 0-2.82l1.17-1.17a2.5 2.5 0 0 1 3.54 0l1.41 1.41a2.5 2.5 0 0 1 0 3.54l-1.17 1.17a2 2 0 0 1-2.82 0')
              );
              icon.appendChild(
                createPath('M13.41 10.59a2 2 0 0 1 0 2.82l-1.17 1.17a2.5 2.5 0 0 1-3.54 0L7.29 13.17a2.5 2.5 0 0 1 0-3.54l1.17-1.17a2 2 0 0 1 2.82 0')
              );

              link.appendChild(icon);

              const srText = doc.createElement('span');
              srText.className = 'visually-hidden';
              srText.textContent = `Abrir tracking: ${displayValue}`;
              link.appendChild(srText);

              td.appendChild(link);
            } else {
              td.classList.add('is-empty');
              td.textContent = '';
            }
          } else if (isStatusColumn) {
            const normalizedStatus = normalizeStatusValue(value);
            if (normalizedStatus) {
              const badge = doc.createElement('span');
              const slug = getStatusBadgeSlug(normalizedStatus);
              badge.className = 'status-badge';
              if (slug) {
                badge.classList.add('status-badge--' + slug);
              }
              badge.textContent = normalizedStatus;
              td.appendChild(badge);
            } else {
              td.classList.add('is-empty');
              td.textContent = '';
            }
          } else {
            td.textContent = typeof value === 'string' ? value : String(value);
          }
          tr.appendChild(td);
        }

        fragment.appendChild(tr);
      });

      refs.tableBody.appendChild(fragment);

      if (rowsToRender.length === 0) {
        setStatus('No hay registros para la vista seleccionada.', 'info');
      } else {
        setStatus('Sincronizado', 'success');
      }
    }

    function getRowDataForIndex(dataIndex) {
      if (!Array.isArray(state.data) || state.data.length <= dataIndex) {
        return null;
      }
      const headers = state.data[0];
      const row = state.data[dataIndex];
      if (!Array.isArray(headers) || !Array.isArray(row)) {
        return null;
      }
      const columnMap = {};
      headers.forEach(function (header, index) {
        const key = getColumnKeyFromHeader(header);
        if (key) {
          columnMap[key] = index;
        }
      });
      const values = {};
      COLUMN_CONFIG.forEach(function (column) {
        const idx = columnMap[column.key];
        if (idx != null) {
          const cellValue = row[idx];
          values[column.key] = cellValue == null ? '' : cellValue;
        } else {
          values[column.key] = '';
        }
      });
      return {
        headers: headers,
        values: values,
        columnMap: columnMap
      };
    }

    function showLoginModal() {
      showBackdrop();
      if (refs.loginModal) {
        refs.loginModal.classList.remove('hidden');
        refs.loginModal.classList.add('is-visible');
      }
      if (refs.loginForm) {
        refs.loginForm.reset();
        if (refs.loginError) {
          refs.loginError.textContent = '';
        }
        const storedUser = loadStoredUser();
        const usernameInput = refs.loginForm.querySelector('input[name="username"]');
        if (usernameInput && storedUser && storedUser.username) {
          usernameInput.value = storedUser.username;
        }
        const tokenInput = refs.loginForm.querySelector('input[name="token"]');
        if (tokenInput) {
          tokenInput.value = state.token || '';
        }
        if (refs.loginTokenField) {
          if (state.token) {
            refs.loginTokenField.classList.add('is-hidden');
          } else {
            refs.loginTokenField.classList.remove('is-hidden');
          }
        }
        const passwordInput = refs.loginForm.querySelector('input[name="password"]');
        if (passwordInput) {
          passwordInput.value = '';
        }
        if (usernameInput) {
          usernameInput.focus();
        }
      }
    }

    function hideLoginModal() {
      if (refs.loginModal) {
        refs.loginModal.classList.remove('is-visible');
        refs.loginModal.classList.add('hidden');
      }
      if (refs.loginError) {
        refs.loginError.textContent = '';
      }
      hideBackdropIfNoModalVisible();
    }

    function openEditModal(dataIndex) {
      if (!refs.editForm || !refs.editModal) {
        return;
      }
      if (!state.currentUser) {
        showLoginModal();
        return;
      }
      const rowData = getRowDataForIndex(dataIndex);
      if (!rowData) {
        setStatus('No fue posible cargar el registro seleccionado.', 'error');
        return;
      }
      const values = rowData.values || {};
      const originalTripValue = values.trip == null ? '' : String(values.trip).trim();
      if (!originalTripValue) {
        setStatus('El registro seleccionado no tiene número de trip.', 'error');
        return;
      }
      state.editingRecord = {
        dataIndex: dataIndex,
        originalTrip: originalTripValue,
        values: values,
        mode: 'edit'
      };
      refs.editForm.reset();
      setEditFormDisabled(false);
      populateEditFormValues(values);
      if (refs.editError) {
        refs.editError.textContent = '';
      }
      setEditModalMode('edit');
      showBackdrop();
      refs.editModal.classList.remove('hidden');
      refs.editModal.classList.add('is-visible');
      const tripInput = refs.editForm.querySelector('[name="trip"]');
      if (tripInput) {
        tripInput.focus();
        if (typeof tripInput.select === 'function') {
          tripInput.select();
        }
      }
    }

    function openCreateModal() {
      if (!refs.editForm || !refs.editModal) {
        return;
      }
      if (!state.currentUser) {
        showLoginModal();
        return;
      }
      const values = COLUMN_CONFIG.reduce(function (acc, column) {
        if (column && column.key) {
          acc[column.key] = '';
        }
        return acc;
      }, {});
      state.editingRecord = {
        dataIndex: null,
        originalTrip: '',
        values: values,
        mode: 'create'
      };
      refs.editForm.reset();
      setEditFormDisabled(false);
      populateEditFormValues(values);
      if (refs.editError) {
        refs.editError.textContent = '';
      }
      setEditModalMode('create');
      showBackdrop();
      refs.editModal.classList.remove('hidden');
      refs.editModal.classList.add('is-visible');
      const tripInput = refs.editForm.querySelector('[name="trip"]');
      if (tripInput) {
        tripInput.focus();
      }
    }

    function closeEditModal() {
      if (refs.editModal) {
        refs.editModal.classList.remove('is-visible');
        refs.editModal.classList.add('hidden');
      }
      if (refs.editError) {
        refs.editError.textContent = '';
      }
      setEditFormDisabled(false);
      if (refs.editForm) {
        refs.editForm.reset();
      }
      state.editingRecord = null;
      hideBackdropIfNoModalVisible();
    }

    function handleTableBodyClick(event) {
      const target = event.target;
      if (!target || typeof target.closest !== 'function') {
        return;
      }
      const trigger = target.closest('[data-action="open-edit"]');
      if (!trigger) {
        return;
      }
      event.preventDefault();
      const rowIndexAttr = trigger.getAttribute('data-row-index');
      const dataIndex = rowIndexAttr == null ? NaN : parseInt(rowIndexAttr, 10);
      if (Number.isNaN(dataIndex)) {
        return;
      }
      openEditModal(dataIndex);
    }

    async function handleEditSubmit(event) {
      event.preventDefault();
      if (!refs.editForm) {
        return;
      }
      if (!state.editingRecord) {
        return;
      }
      const formData = new global.FormData(refs.editForm);
      const tripValue = String(formData.get('trip') || '').trim();
      const ejecutivoValue = String(formData.get('ejecutivo') || '').trim();
      if (!tripValue || !/^\d+$/.test(tripValue)) {
        if (refs.editError) {
          refs.editError.textContent = 'Ingresa un número de trip válido.';
        }
        const tripInput = refs.editForm.querySelector('[name="trip"]');
        if (tripInput) {
          tripInput.focus();
        }
        return;
      }
      if (!ejecutivoValue) {
        if (refs.editError) {
          refs.editError.textContent = 'El campo Ejecutivo es obligatorio.';
        }
        const ejecutivoInput = refs.editForm.querySelector('[name="ejecutivo"]');
        if (ejecutivoInput) {
          ejecutivoInput.focus();
        }
        return;
      }

      const mode = state.editingRecord.mode === 'create' ? 'create' : 'edit';
      if (mode === 'edit' && !state.editingRecord.originalTrip) {
        if (refs.editError) {
          refs.editError.textContent = 'No se encontró el trip original del registro.';
        }
        return;
      }
      if (!state.config || !state.config.API_BASE) {
        const message = 'Falta configurar la URL del Apps Script.';
        if (refs.editError) {
          refs.editError.textContent = message;
        }
        setStatus(message, 'error');
        return;
      }
      if (!state.token) {
        const message = 'Sesión expirada. Inicia sesión nuevamente.';
        if (refs.editError) {
          refs.editError.textContent = message;
        }
        setStatus(message, 'error');
        closeEditModal();
        showLoginModal();
        return;
      }

      function getTrimmed(name) {
        const value = formData.get(name);
        return value == null ? '' : String(value).trim();
      }

      const payload = {
        action: mode === 'create' ? 'add' : 'update',
        trip: tripValue,
        ejecutivo: ejecutivoValue,
        caja: getTrimmed('caja'),
        referencia: getTrimmed('referencia'),
        cliente: getTrimmed('cliente'),
        destino: getTrimmed('destino'),
        estatus: getTrimmed('estatus'),
        segmento: getTrimmed('segmento'),
        trmx: getTrimmed('trmx'),
        trusa: getTrimmed('trusa'),
        citaCarga: toApiDateValue(formData.get('citaCarga')),
        llegadaCarga: toApiDateValue(formData.get('llegadaCarga')),
        citaEntrega: toApiDateValue(formData.get('citaEntrega')),
        llegadaEntrega: toApiDateValue(formData.get('llegadaEntrega')),
        comentarios: (function () {
          const value = formData.get('comentarios');
          return value == null ? '' : String(value);
        })(),
        docs: getTrimmed('docs'),
        tracking: getTrimmed('tracking')
      };

      if (mode === 'edit') {
        payload.originalTrip = state.editingRecord.originalTrip;
      }

      if (refs.editError) {
        refs.editError.textContent = '';
      }
      setEditFormDisabled(true);
      setStatus('Guardando cambios…', 'info');

      try {
        await submitRecordRequest(state.config.API_BASE, state.token, payload);
        closeEditModal();
        await loadData();
        if (!refs.status || !refs.status.classList.contains('is-error')) {
          const successMessage =
            mode === 'create'
              ? 'Registro agregado correctamente.'
              : 'Registro actualizado correctamente.';
          setStatus(successMessage, 'success');
        }
      } catch (err) {
        const fallbackMessage =
          mode === 'create' ? 'Error al agregar el registro.' : 'Error al actualizar el registro.';
        const message = err && err.message ? err.message : fallbackMessage;
        if (err && err.status === 401) {
          state.token = '';
          setStoredValue(STORAGE_TOKEN_KEY, null);
          state.currentUser = null;
          updateUserBadge();
          clearTable();
          updateLastUpdated(null);
          closeEditModal();
          setStatus('Sesión expirada. Inicia sesión nuevamente.', 'error');
          showLoginModal();
          return;
        }
        if (refs.editError) {
          refs.editError.textContent = message;
        }
        setStatus(message, 'error');
      } finally {
        setEditFormDisabled(false);
      }
    }

    function handleLoginSubmit(event) {
      event.preventDefault();
      if (!refs.loginForm) return;
      const formData = new global.FormData(refs.loginForm);
      const username = String(formData.get('username') || '').trim();
      const password = String(formData.get('password') || '');
      const tokenInput = String(formData.get('token') || '').trim();

      if (!username || !password) {
        if (refs.loginError) {
          refs.loginError.textContent = 'Ingresa usuario y contraseña.';
        }
        return;
      }

      let tokenToUse = state.token;
      if (!tokenToUse && tokenInput) {
        tokenToUse = tokenInput;
      }

      if (!tokenToUse) {
        if (refs.loginError) {
          refs.loginError.textContent = 'Ingresa el token proporcionado por el Apps Script.';
        }
        return;
      }

      const match = state.users.find(function (user) {
        return user.id === username.toLowerCase();
      }) || (state.users.length === 0 && username.toLowerCase() === DEFAULT_USER.id ? DEFAULT_USER : null);

      if (!match || match.password !== password) {
        if (refs.loginError) {
          refs.loginError.textContent = 'Credenciales inválidas. Verifica usuario y contraseña.';
        }
        return;
      }

      state.currentUser = match;
      state.token = tokenToUse;
      setStoredValue(STORAGE_TOKEN_KEY, tokenToUse);
      setStoredValue(STORAGE_USER_KEY, { username: match.username, displayName: match.displayName });
      updateUserBadge();
      hideLoginModal();
      setStatus(`Sesión iniciada como ${match.displayName}.`, 'success');
      loadData();
    }

    function handleLogout() {
      closeEditModal();
      state.currentUser = null;
      resetFilters();
      updateUserBadge();
      setStatus('Sesión cerrada. Vuelve a iniciar sesión para ver la hoja.', 'info');
      clearTable();
      updateLastUpdated(null);
      showLoginModal();
    }

    function handleChangeToken() {
      const promptValue = global.prompt('Introduce el token actualizado del Apps Script:', state.token || '');
      if (promptValue == null) {
        return;
      }
      const trimmed = String(promptValue).trim();
      if (!trimmed) {
        setStatus('El token no puede estar vacío.', 'error');
        return;
      }
      state.token = trimmed;
      setStoredValue(STORAGE_TOKEN_KEY, trimmed);
      if (state.currentUser) {
        setStatus('Token actualizado. Sincronizando datos…', 'info');
        loadData();
      } else {
        setStatus('Token actualizado. Inicia sesión para continuar.', 'success');
      }
    }

    async function loadData() {
      if (!state.currentUser) {
        showLoginModal();
        return;
      }
      if (!state.token) {
        setStatus('No se encontró el token. Actualízalo para continuar.', 'error');
        showLoginModal();
        return;
      }
      toggleLoading(true);
      setStatus('Cargando datos desde la hoja…', 'info');
      try {
        const rows = await fetchSheetData(state.config.API_BASE, state.token);
        state.data = rows;
        renderTable();
        updateLastUpdated(new Date());
      } catch (err) {
        const message = err && err.message ? err.message : 'Error al cargar los datos.';
        setStatus(message, 'error');
        if (err && err.status === 401) {
          state.token = '';
          setStoredValue(STORAGE_TOKEN_KEY, null);
          showLoginModal();
        }
      } finally {
        toggleLoading(false);
      }
    }

    async function bootstrap() {
      const secureConfig = await fetchSecureConfig(state.config.SECURE_CONFIG_URL);
      state.secureConfigLoaded = true;
      const configToken = secureConfig.API_TOKEN || secureConfig.apiToken || null;
      const storedToken = getStoredValue(STORAGE_TOKEN_KEY);
      state.token = (configToken || storedToken || '').trim();
      if (state.token) {
        setStoredValue(STORAGE_TOKEN_KEY, state.token);
      }

      const normalizedUsers = normalizeUsers(secureConfig);
      if (normalizedUsers.length > 0) {
        state.users = normalizedUsers;
      } else {
        state.users = [DEFAULT_USER];
      }

      showLoginModal();
    }

    if (refs.loginForm) {
      refs.loginForm.addEventListener('submit', handleLoginSubmit);
    }
    if (refs.tableBody) {
      refs.tableBody.addEventListener('click', handleTableBodyClick);
    }
    if (refs.editForm) {
      refs.editForm.addEventListener('submit', handleEditSubmit);
    }
    if (refs.cancelEditButton) {
      refs.cancelEditButton.addEventListener('click', function () {
        closeEditModal();
      });
    }
    if (refs.refreshButton) {
      refs.refreshButton.addEventListener('click', function () {
        loadData();
      });
    }
    if (refs.newRecordButton) {
      refs.newRecordButton.addEventListener('click', function () {
        openCreateModal();
      });
    }
    if (refs.logoutButton) {
      refs.logoutButton.addEventListener('click', handleLogout);
    }
    if (refs.changeTokenButton) {
      refs.changeTokenButton.addEventListener('click', handleChangeToken);
    }
    if (refs.themeSwitch) {
      refs.themeSwitch.addEventListener('change', handleThemeSwitchChange);
    }
    if (refs.filterSearchInput) {
      refs.filterSearchInput.addEventListener('input', handleFilterSearchInput);
    }
    if (refs.dateToggleButton) {
      refs.dateToggleButton.addEventListener('click', handleDateToggle);
    }
    if (refs.datePrevButton) {
      refs.datePrevButton.addEventListener('click', handleDatePrev);
    }
    if (refs.dateNextButton) {
      refs.dateNextButton.addEventListener('click', handleDateNext);
    }
    if (refs.dateStartInput) {
      refs.dateStartInput.addEventListener('change', handleDateInputChange);
    }
    if (refs.dateEndInput) {
      refs.dateEndInput.addEventListener('change', handleDateInputChange);
    }
    if (refs.dateClearButton) {
      refs.dateClearButton.addEventListener('click', handleDateClear);
    }
    if (refs.statusToggleButton) {
      refs.statusToggleButton.addEventListener('click', handleStatusToggle);
    }
    if (refs.statusList) {
      refs.statusList.addEventListener('click', handleStatusListClick);
    }
    if (refs.viewMenu) {
      refs.viewMenu.addEventListener('click', handleViewMenuClick);
    }

    doc.addEventListener('click', handleDocumentClick);
    doc.addEventListener('keydown', handleDocumentKeydown);

    renderViewMenu();

    bootstrap();

    if ('serviceWorker' in global.navigator) {
      global.navigator.serviceWorker.register('./sw.js').catch(function () {
        // ignore registration errors
      });
    }
  }

  if (typeof window !== 'undefined' && window.document) {
    window.addEventListener('DOMContentLoaded', initApp);
  }

  const exportsObject = {
    fmtDate: fmtDate,
    DEFAULT_LOCALE: DEFAULT_LOCALE,
    formatHeaderLabel: formatHeaderLabel
  };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = exportsObject;
  } else {
    global.App = Object.assign({}, global.App || {}, exportsObject);
  }
})(typeof window !== 'undefined' ? window : globalThis);
