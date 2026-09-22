(function () {
  'use strict';

  var CFG = window.SITE_CONFIG || {};
  var DEFAULTS = window.DEFAULT_CONTENT || {};
  var SHEET_ID = String(CFG.SHEET_ID || '').trim();
  var TIMEOUT_MS = Number(CFG.TIMEOUT_MS) > 0 ? Number(CFG.TIMEOUT_MS) : 6000;

  var SHEETS = {
    settings: { name: 'Настройки', required: ['ключ', 'значение'] },
    trainings: { name: 'Тренировки', required: ['название', 'описание'] },
    schedule: { name: 'Расписание', required: ['день', 'время'] },
    venues: { name: 'Площадки', required: ['название', 'адрес'] },
    promos: { name: 'Акции', required: ['заголовок', 'текст'] },
    tournaments: { name: 'Турниры', required: ['название'] },
    coaches: { name: 'Тренер', required: ['имя'] },
    partners: { name: 'Партнёры', required: ['название'] },
    faq: { name: 'FAQ', required: ['вопрос', 'ответ'] },
    socials: { name: 'Соцсети', required: ['название', 'ссылка'] }
  };

  var FALLBACK_TELEGRAM = 'https://t.me/furiousrackets';

  var DAY_SHORT = {
    'понедельник': 'пн',
    'вторник': 'вт',
    'среда': 'ср',
    'четверг': 'чт',
    'пятница': 'пт',
    'суббота': 'сб',
    'воскресенье': 'вс'
  };
  var DAY_ORDER = ['пн', 'вт', 'ср', 'чт', 'пт', 'сб', 'вс'];

  /* ---------- утилиты ---------- */

  function q(sel) { return document.querySelector(sel); }
  function qa(sel) { return Array.prototype.slice.call(document.querySelectorAll(sel)); }

  var ESCAPES = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };

  function esc(value) {
    return String(value == null ? '' : value).replace(/[&<>"']/g, function (ch) {
      return ESCAPES[ch];
    });
  }

  function escMultiline(value) {
    return esc(value).replace(/\r\n|\r|\n/g, '<br>');
  }

  var LINKS = {};

  function richText(value) {
    var html = escMultiline(value);
    return html.replace(/\{(телефон|тренер|группа)\}/g, function (match, name) {
      return LINKS[name] || '';
    });
  }

  function handleFromUrl(url) {
    var match = /^https?:\/\/[^/]+\/(?:joinchat\/)?([^/?#]+)/i.exec(String(url || ''));
    return match ? '@' + match[1] : '';
  }

  function buildLinks(settings) {
    LINKS = {};

    var phone = settings.phone || '';
    var phoneHref = telHref(phone);
    if (phoneHref) {
      LINKS['телефон'] = '<a class="link-accent" href="' + phoneHref + '">' + esc(phone) + '</a>';
    }

    var coach = safeUrl(settings.coach_telegram);
    if (coach) {
      LINKS['тренер'] = '<a class="link-soft" href="' + coach + '" target="_blank" rel="noopener">' +
        esc(handleFromUrl(settings.coach_telegram) || 'в Telegram') + '</a>';
    }

    var group = safeUrl(settings.group_telegram) || FALLBACK_TELEGRAM;
    LINKS['группа'] = '<a class="link-accent" href="' + group + '" target="_blank" rel="noopener">' +
      esc(handleFromUrl(group) || 'в Telegram') + '</a>';
  }

  function stripControls(value) {
    var text = String(value == null ? '' : value);
    var out = '';
    for (var i = 0; i < text.length; i += 1) {
      var code = text.charCodeAt(i);
      if (code > 31 && code !== 127) out += text.charAt(i);
    }
    return out;
  }

  function safeUrl(value) {
    var raw = stripControls(value).trim();
    if (!raw) return '';
    if (/^\/\//.test(raw)) return '';
    if (/^(https?:|tel:|mailto:)/i.test(raw)) return esc(raw);
    if (/^[a-z][a-z0-9+.\-]*:/i.test(raw)) return '';
    return esc(raw);
  }

  function telHref(phone) {
    var digits = String(phone || '').replace(/[^\d+]/g, '');
    return digits ? 'tel:' + esc(digits) : '';
  }

  function normalizeHeader(value) {
    return String(value == null ? '' : value)
      .replace(/﻿/g, '')
      .trim()
      .toLowerCase()
      .replace(/ё/g, 'е')
      .replace(/\s+/g, ' ');
  }

  function cell(row, key) {
    var value = row ? row[key] : '';
    return String(value == null ? '' : value).trim();
  }

  /* ---------- CSV ---------- */

  function parseCSV(text) {
    var input = String(text == null ? '' : text).replace(/^﻿/, '');
    var rows = [];
    var row = [];
    var field = '';
    var inQuotes = false;
    var i = 0;

    while (i < input.length) {
      var ch = input.charAt(i);

      if (inQuotes) {
        if (ch === '"') {
          if (input.charAt(i + 1) === '"') {
            field += '"';
            i += 2;
            continue;
          }
          inQuotes = false;
          i += 1;
          continue;
        }
        field += ch;
        i += 1;
        continue;
      }

      if (ch === '"') {
        inQuotes = true;
        i += 1;
        continue;
      }
      if (ch === ',') {
        row.push(field);
        field = '';
        i += 1;
        continue;
      }
      if (ch === '\r' || ch === '\n') {
        row.push(field);
        rows.push(row);
        row = [];
        field = '';
        i += (ch === '\r' && input.charAt(i + 1) === '\n') ? 2 : 1;
        continue;
      }
      field += ch;
      i += 1;
    }

    row.push(field);
    rows.push(row);
    return rows;
  }

  function toObjects(matrix) {
    if (!matrix || !matrix.length) return null;
    var headers = matrix[0].map(normalizeHeader);
    var result = [];

    for (var r = 1; r < matrix.length; r += 1) {
      var source = matrix[r];
      var item = {};
      var hasValue = false;

      for (var c = 0; c < headers.length; c += 1) {
        if (!headers[c]) continue;
        var value = String(source[c] == null ? '' : source[c]).trim();
        item[headers[c]] = value;
        if (value) hasValue = true;
      }
      if (hasValue) result.push(item);
    }
    return { headers: headers, rows: result };
  }

  function hasRequiredHeaders(table, required) {
    if (!table) return false;
    for (var i = 0; i < required.length; i += 1) {
      if (table.headers.indexOf(required[i]) === -1) return false;
    }
    return true;
  }

  /* ---------- фильтры строк ---------- */

  var HIDDEN_WORDS = ['нет', 'no', '0', 'false'];

  function isShown(row) {
    var flag = cell(row, 'показывать').toLowerCase().replace(/ё/g, 'е');
    if (!flag) return true;
    return HIDDEN_WORDS.indexOf(flag) === -1;
  }

  function parseDate(value) {
    var text = String(value || '').trim();
    var match = /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/.exec(text);
    var year, month, day;

    if (match) {
      day = Number(match[1]);
      month = Number(match[2]);
      year = Number(match[3]);
    } else {
      match = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(text);
      if (!match) return null;
      year = Number(match[1]);
      month = Number(match[2]);
      day = Number(match[3]);
    }

    var date = new Date(year, month - 1, day);
    if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
      return null;
    }
    return date;
  }

  function isActual(row) {
    var raw = cell(row, 'до');
    if (!raw) return true;
    var until = parseDate(raw);
    if (!until) return true;
    var now = new Date();
    var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    return today.getTime() <= until.getTime();
  }

  function visibleRows(rows) {
    return (rows || []).filter(function (row) {
      return isShown(row) && isActual(row);
    });
  }

  /* ---------- загрузка ---------- */

  function sheetUrl(name) {
    return 'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(SHEET_ID) +
      '/gviz/tq?tqx=out:csv&headers=1&sheet=' + encodeURIComponent(name);
  }

  function fetchSheet(name) {
    if (typeof fetch !== 'function') return Promise.reject(new Error('no fetch'));

    var controller = typeof AbortController === 'function' ? new AbortController() : null;
    var options = controller ? { signal: controller.signal } : {};
    var timer;

    var request = fetch(sheetUrl(name), options).then(function (response) {
      if (!response.ok) throw new Error('HTTP ' + response.status);
      return response.text();
    });

    var timeout = new Promise(function (_, reject) {
      timer = setTimeout(function () {
        if (controller) controller.abort();
        reject(new Error('timeout'));
      }, TIMEOUT_MS);
    });

    return Promise.race([request, timeout]).then(function (text) {
      clearTimeout(timer);
      return text;
    }, function (error) {
      clearTimeout(timer);
      throw error;
    });
  }

  function defaultRows(sheetName) {
    var rows = DEFAULTS[sheetName];
    return Array.isArray(rows) ? rows.slice() : [];
  }

  function loadSheet(key) {
    var meta = SHEETS[key];
    if (!SHEET_ID) return Promise.resolve(defaultRows(meta.name));

    return fetchSheet(meta.name).then(function (text) {
      var table = toObjects(parseCSV(text));
      if (!hasRequiredHeaders(table, meta.required)) return defaultRows(meta.name);
      return table.rows;
    }).catch(function () {
      return defaultRows(meta.name);
    });
  }

  function loadContent() {
    var keys = Object.keys(SHEETS);
    return Promise.all(keys.map(loadSheet)).then(function (results) {
      var data = {};
      keys.forEach(function (key, index) {
        data[key] = results[index];
      });
      return data;
    });
  }

  /* ---------- настройки ---------- */

  function toSettings(rows) {
    var map = {};
    (rows || []).forEach(function (row) {
      var key = cell(row, 'ключ').toLowerCase();
      if (key) map[key] = cell(row, 'значение');
    });
    return map;
  }

  function applySettings(settings) {
    buildLinks(settings);

    var telegram = safeUrl(settings.group_telegram) || FALLBACK_TELEGRAM;
    var phone = settings.phone || '';
    var phoneLink = telHref(phone);

    qa('[data-tg-link]').forEach(function (node) {
      node.setAttribute('href', telegram);
    });

    qa('[data-phone-link]').forEach(function (node) {
      if (phoneLink) {
        node.setAttribute('href', phoneLink);
        node.hidden = false;
      } else {
        node.hidden = true;
      }
    });
    qa('[data-phone-text]').forEach(function (node) {
      node.textContent = phone;
    });

    qa('[data-phone-name]').forEach(function (node) {
      if (settings.phone_name) {
        node.innerHTML = escMultiline(settings.phone_name);
        node.hidden = false;
      } else {
        node.hidden = true;
      }
    });

    setText('[data-hero-title]', settings.hero_title);
    setText('[data-hero-slogan]', settings.hero_slogan);
    setText('[data-hero-text]', settings.hero_text);

    var intro = q('[data-tournaments-intro]');
    if (intro) {
      if (settings.tournaments_intro) {
        intro.innerHTML = escMultiline(settings.tournaments_intro);
        intro.hidden = false;
      } else {
        intro.hidden = true;
      }
    }
  }

  function setText(selector, value) {
    var node = q(selector);
    if (node && value) node.innerHTML = escMultiline(value);
  }

  /* ---------- отрисовка ---------- */

  function reveal(sectionId, isVisible) {
    var section = document.getElementById(sectionId);
    if (section) section.hidden = !isVisible;
    var link = q('[data-nav="' + sectionId + '"]');
    if (link) link.hidden = !isVisible;
  }

  function renderTrainings(rows) {
    var items = visibleRows(rows);
    var host = q('[data-trainings]');
    if (!host) return;

    host.innerHTML = items.map(function (row) {
      return '<article class="training">' +
        '<h3 class="training__title">' + esc(cell(row, 'название')) + '</h3>' +
        '<p class="training__text">' + richText(cell(row, 'описание')) + '</p>' +
        '</article>';
    }).join('');

    reveal('trainings', items.length > 0);
  }

  function formatPrice(value) {
    var price = String(value == null ? '' : value).trim();
    if (!price) return '';
    if (/^\d+([.,]\d+)?$/.test(price)) return price + ' ₽';
    return price;
  }

  function levelBadges(value) {
    var levels = String(value || '').split(',').map(function (part) {
      return part.trim();
    }).filter(Boolean);

    if (!levels.length) {
      return '<span class="level level--any">все уровни</span>';
    }
    return levels.map(function (level) {
      return '<span class="level">' + esc(level) + '</span>';
    }).join('');
  }

  function dayShort(day) {
    var key = String(day || '').trim().toLowerCase().replace(/ё/g, 'е');
    if (DAY_SHORT[key]) return DAY_SHORT[key];
    return key.slice(0, 2);
  }

  function renderSchedule(rows, venueIndex) {
    var items = visibleRows(rows);
    var host = q('[data-schedule]');
    if (!host) return items;

    host.innerHTML = items.map(function (row) {
      var type = cell(row, 'тип');
      var isPlay = type.toLowerCase().indexOf('игров') === 0;
      var venueKey = cell(row, 'площадка');
      var venue = venueIndex.byName[venueKey.toLowerCase()];
      var venueTitle = venue ? (venue.fullName || venue.name) : venueKey;
      var venueHtml = esc(venueTitle);

      if (venue && venue.anchor) {
        venueHtml = '<a class="row__venue-link" href="#' + venue.anchor + '">' + esc(venueTitle) + '</a>';
      }

      var price = formatPrice(cell(row, 'цена'));

      return '<article class="row">' +
        '<div class="row__day"><span class="row__day-in">' + esc(cell(row, 'день')) + '</span></div>' +
        '<div class="row__time">' + esc(cell(row, 'время')) + '</div>' +
        '<div class="row__meta">' +
          (type ? '<span class="type' + (isPlay ? ' type--play' : '') + '">' + esc(type) + '</span>' : '') +
          '<span class="levels">' + levelBadges(cell(row, 'уровни')) + '</span>' +
        '</div>' +
        '<div class="row__venue">' + venueHtml + '</div>' +
        '<div class="row__price">' + (price ? '<span class="price">' + esc(price) + '</span>' : '') + '</div>' +
        '</article>';
    }).join('');

    reveal('schedule', items.length > 0);
    return items;
  }

  function buildVenueIndex(rows) {
    var all = (rows || []).slice();
    var shown = visibleRows(all);
    var byName = {};

    all.forEach(function (row) {
      var name = cell(row, 'название');
      if (!name) return;
      byName[name.toLowerCase()] = {
        name: name,
        fullName: cell(row, 'полное название'),
        anchor: ''
      };
    });

    shown.forEach(function (row, index) {
      var name = cell(row, 'название');
      var entry = byName[name.toLowerCase()];
      if (entry) entry.anchor = 'venue-' + (index + 1);
    });

    return { byName: byName, shown: shown };
  }

  function renderVenues(venueIndex, scheduleRows) {
    var host = q('[data-venues]');
    if (!host) return;

    var daysByVenue = {};
    visibleRows(scheduleRows).forEach(function (row) {
      var key = cell(row, 'площадка').toLowerCase();
      var short = dayShort(cell(row, 'день'));
      if (!key || !short) return;
      if (!daysByVenue[key]) daysByVenue[key] = [];
      if (daysByVenue[key].indexOf(short) === -1) daysByVenue[key].push(short);
    });

    host.innerHTML = venueIndex.shown.map(function (row, index) {
      var name = cell(row, 'название');
      var fullName = cell(row, 'полное название') || name;
      var address = cell(row, 'адрес');
      var map = safeUrl(cell(row, 'ссылка на карту'));

      if (!map && address) {
        map = esc('https://yandex.ru/maps/?text=' + encodeURIComponent(address));
      }

      var days = (daysByVenue[name.toLowerCase()] || []).slice().sort(function (a, b) {
        return DAY_ORDER.indexOf(a) - DAY_ORDER.indexOf(b);
      });

      return '<article class="venue" id="venue-' + (index + 1) + '">' +
        '<h3 class="venue__title">' + esc(fullName) + '</h3>' +
        (address ? '<p class="venue__address">' + escMultiline(address) + '</p>' : '') +
        '<div class="venue__foot">' +
          '<p class="venue__days">' + days.map(function (day) {
            return '<span class="day-chip">' + esc(day) + '</span>';
          }).join('') + '</p>' +
          (map ? '<a class="venue__map" href="' + map + '" target="_blank" rel="noopener">Показать на карте</a>' : '') +
        '</div>' +
        '</article>';
    }).join('');

    reveal('venues', venueIndex.shown.length > 0);
  }

  function renderPromos(rows, telegram) {
    var items = visibleRows(rows);
    var host = q('[data-promos]');
    if (!host) return;

    host.innerHTML = items.map(function (row) {
      var image = safeUrl(cell(row, 'картинка'));
      var label = cell(row, 'текст кнопки');
      var link = safeUrl(cell(row, 'ссылка кнопки')) || telegram;
      var title = cell(row, 'заголовок');

      return '<article class="promo">' +
        (image ? '<div class="promo__media"><img src="' + image + '" alt="' + esc(title) + '" loading="lazy"></div>' : '') +
        '<div class="promo__body">' +
          '<h3 class="promo__title">' + escMultiline(title) + '</h3>' +
          '<p class="promo__text">' + richText(cell(row, 'текст')) + '</p>' +
          (label && link
            ? '<a class="btn btn--accent" href="' + link + '" target="_blank" rel="noopener">' + esc(label) + '</a>'
            : '') +
        '</div>' +
        '</article>';
    }).join('');

    reveal('promos', items.length > 0);
  }

  function renderTournaments(rows, venueIndex) {
    var items = visibleRows(rows);
    var host = q('[data-tournaments]');
    if (!host) return;

    host.innerHTML = items.map(function (row) {
      var poster = safeUrl(cell(row, 'афиша'));
      var link = safeUrl(cell(row, 'ссылка'));
      var title = cell(row, 'название');
      var placeKey = cell(row, 'место');
      var venue = venueIndex.byName[placeKey.toLowerCase()];
      var place = venue ? (venue.fullName || venue.name) : placeKey;

      return '<article class="tournament">' +
        (poster ? '<div class="tournament__media"><img src="' + poster + '" alt="' + esc(title) + '" loading="lazy"></div>' : '') +
        '<div class="tournament__body">' +
          '<h3 class="tournament__title">' + escMultiline(title) + '</h3>' +
          '<p class="tournament__when">' +
            esc(cell(row, 'даты')) +
            (place ? '<span class="tournament__where">' + esc(place) + '</span>' : '') +
          '</p>' +
          '<p class="tournament__text">' + richText(cell(row, 'текст')) + '</p>' +
          (link ? '<a class="btn btn--ghost" href="' + link + '" target="_blank" rel="noopener">Подробнее</a>' : '') +
        '</div>' +
        '</article>';
    }).join('');

    reveal('tournaments', items.length > 0);
  }

  function renderCoaches(rows) {
    var items = visibleRows(rows).filter(function (row) {
      return cell(row, 'имя');
    });
    var host = q('[data-coaches]');
    if (!host) return;

    host.innerHTML = items.map(function (row) {
      var photo = safeUrl(cell(row, 'фото'));
      var telegram = safeUrl(cell(row, 'telegram'));
      var name = cell(row, 'имя');

      return '<article class="coach">' +
        (photo ? '<div class="coach__media"><img src="' + photo + '" alt="' + esc(name) + '" loading="lazy"></div>' : '') +
        '<div class="coach__body">' +
          '<h3 class="coach__name">' + esc(name) + '</h3>' +
          '<p class="coach__about">' + escMultiline(cell(row, 'о себе')) + '</p>' +
          (telegram ? '<a class="link-accent" href="' + telegram + '" target="_blank" rel="noopener">Написать в Telegram</a>' : '') +
        '</div>' +
        '</article>';
    }).join('');

    reveal('coach', items.length > 0);
  }

  function renderPartners(rows) {
    var items = visibleRows(rows);
    var host = q('[data-partners]');
    if (!host) return;

    host.innerHTML = items.map(function (row) {
      var name = cell(row, 'название');
      var logo = safeUrl(cell(row, 'логотип'));
      var link = safeUrl(cell(row, 'ссылка'));
      var description = cell(row, 'описание');

      var inner = (logo
        ? '<img class="partner__logo" src="' + logo + '" alt="' + esc(name) + '" loading="lazy">'
        : '<span class="partner__name">' + esc(name) + '</span>') +
        (description ? '<span class="partner__desc">' + escMultiline(description) + '</span>' : '');

      return link
        ? '<a class="partner" href="' + link + '" target="_blank" rel="noopener">' + inner + '</a>'
        : '<div class="partner">' + inner + '</div>';
    }).join('');

    reveal('partners', items.length > 0);
  }

  function renderSocials(rows) {
    var host = q('[data-socials]');
    if (!host) return;

    var items = visibleRows(rows).filter(function (row) {
      return cell(row, 'название') && safeUrl(cell(row, 'ссылка'));
    });

    host.innerHTML = items.map(function (row) {
      return '<li class="socials__item">' +
        '<a class="socials__link" href="' + safeUrl(cell(row, 'ссылка')) + '" target="_blank" rel="noopener">' +
        esc(cell(row, 'название')) + '</a></li>';
    }).join('');

    host.hidden = items.length === 0;
  }

  function renderFaq(rows) {
    var items = visibleRows(rows);
    var host = q('[data-faq]');
    if (!host) return;

    host.innerHTML = items.map(function (row) {
      return '<details class="qa">' +
        '<summary class="qa__q">' + esc(cell(row, 'вопрос')) + '</summary>' +
        '<div class="qa__a">' + richText(cell(row, 'ответ')) + '</div>' +
        '</details>';
    }).join('');

    reveal('faq', items.length > 0);
  }

  /* ---------- служебное ---------- */

  function createLoader() {
    var root = document.documentElement;
    var startedAt = Date.now();

    function hide() {
      root.className = root.className.replace(/\bis-loading\b/, '');
    }

    return {
      done: function () {
        if (root.className.indexOf('is-loading') === -1) return;
        var elapsed = Date.now() - startedAt;
        if (elapsed < 400) {
          setTimeout(hide, 400 - elapsed);
        } else {
          hide();
        }
      }
    };
  }

  function setYear() {
    var node = q('[data-year]');
    if (node) node.textContent = String(new Date().getFullYear());
  }

  function setupStickyCta() {
    var cta = q('[data-sticky-cta]');
    var hero = q('.hero');
    if (!cta || !hero) return;

    if (typeof IntersectionObserver !== 'function') {
      cta.classList.add('is-visible');
      return;
    }

    var observer = new IntersectionObserver(function (entries) {
      entries.forEach(function (entry) {
        cta.classList.toggle('is-visible', !entry.isIntersecting);
      });
    }, { rootMargin: '-80px 0px 0px 0px' });

    observer.observe(hero);
  }

  function applyHash() {
    var hash = String(location.hash || '');
    if (hash.length < 2) return;

    var target = null;
    try {
      target = document.querySelector(hash);
    } catch (error) {
      return;
    }
    if (!target || target.hidden) return;

    var scroll = function () {
      try {
        target.scrollIntoView({ behavior: 'instant', block: 'start' });
      } catch (error) {
        target.scrollIntoView();
      }
    };

    var fonts = document.fonts;
    if (!fonts || !fonts.ready || typeof fonts.ready.then !== 'function') {
      scroll();
      return;
    }

    var done = false;
    var run = function () {
      if (done) return;
      done = true;
      scroll();
    };
    fonts.ready.then(run, run);
    setTimeout(run, 800);
  }

  function render(data) {
    var settings = toSettings(data.settings);
    applySettings(settings);

    var telegram = safeUrl(settings.group_telegram) || FALLBACK_TELEGRAM;
    var venueIndex = buildVenueIndex(data.venues);

    renderTrainings(data.trainings);
    renderSchedule(data.schedule, venueIndex);
    renderVenues(venueIndex, data.schedule);
    renderPromos(data.promos, telegram);
    renderCoaches(data.coaches);
    renderTournaments(data.tournaments, venueIndex);
    renderPartners(data.partners);
    renderFaq(data.faq);
    renderSocials(data.socials);

    applyHash();
  }

  function start() {
    setYear();
    setupStickyCta();

    var loader = createLoader();

    loadContent().then(render).catch(function () {
      var fallback = {};
      Object.keys(SHEETS).forEach(function (key) {
        fallback[key] = defaultRows(SHEETS[key].name);
      });
      render(fallback);
    }).then(loader.done, loader.done);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', start);
  } else {
    start();
  }
})();
