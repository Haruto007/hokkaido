const STORE_KEY = 'hokkaido_trip_planner_v1';

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('北海道旅ノート')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getTripData() {
  return readStore_();
}

function saveExpense(expense) {
  const data = readStore_();
  const now = new Date().toISOString();
  const normalized = {
    id: expense.id || Utilities.getUuid(),
    date: expense.date || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd'),
    payer: String(expense.payer || '').trim(),
    item: String(expense.item || '').trim(),
    amount: Number(expense.amount || 0),
    splitWith: normalizeNames_(expense.splitWith),
    memo: String(expense.memo || '').trim(),
    createdAt: expense.createdAt || now,
    updatedAt: now,
  };

  if (!normalized.payer || !normalized.item || normalized.amount <= 0) {
    throw new Error('支払い者・内容・金額を入力してください。');
  }

  if (!normalized.splitWith.length) {
    normalized.splitWith = uniqueNames_([normalized.payer].concat(data.members));
  }

  const index = data.expenses.findIndex((item) => item.id === normalized.id);
  if (index >= 0) {
    data.expenses[index] = normalized;
  } else {
    data.expenses.unshift(normalized);
  }

  data.members = uniqueNames_(data.members.concat([normalized.payer], normalized.splitWith));
  writeStore_(data);
  return data;
}

function deleteExpense(id) {
  const data = readStore_();
  data.expenses = data.expenses.filter((expense) => expense.id !== id);
  writeStore_(data);
  return data;
}

function savePlace(place) {
  const data = readStore_();
  const now = new Date().toISOString();
  const normalized = {
    id: place.id || Utilities.getUuid(),
    name: String(place.name || '').trim(),
    address: String(place.address || '').trim(),
    note: String(place.note || '').trim(),
    day: String(place.day || '').trim(),
    status: place.status || '行きたい',
    order: Number.isFinite(Number(place.order)) ? Number(place.order) : data.places.length,
    createdAt: place.createdAt || now,
    updatedAt: now,
  };

  if (!normalized.name) {
    throw new Error('場所の名前を入力してください。');
  }

  const index = data.places.findIndex((item) => item.id === normalized.id);
  if (index >= 0) {
    data.places[index] = normalized;
  } else {
    data.places.push(normalized);
  }

  data.places = normalizePlaceOrder_(data.places);
  writeStore_(data);
  return data;
}

function deletePlace(id) {
  const data = readStore_();
  data.places = normalizePlaceOrder_(data.places.filter((place) => place.id !== id));
  writeStore_(data);
  return data;
}

function updatePlaceOrder(placeIds) {
  const data = readStore_();
  const orderMap = {};
  placeIds.forEach((id, index) => {
    orderMap[id] = index;
  });

  data.places = normalizePlaceOrder_(data.places.map((place) => {
    if (Object.prototype.hasOwnProperty.call(orderMap, place.id)) {
      return Object.assign({}, place, { order: orderMap[place.id] });
    }
    return place;
  }));

  writeStore_(data);
  return data;
}

function updateMembers(members) {
  const data = readStore_();
  data.members = normalizeNames_(members);
  writeStore_(data);
  return data;
}

function resetTripData() {
  const data = defaultStore_();
  writeStore_(data);
  return data;
}

function readStore_() {
  const raw = PropertiesService.getScriptProperties().getProperty(STORE_KEY);
  if (!raw) {
    return defaultStore_();
  }

  try {
    const parsed = JSON.parse(raw);
    return {
      members: normalizeNames_(parsed.members || []),
      expenses: Array.isArray(parsed.expenses) ? parsed.expenses : [],
      places: normalizePlaceOrder_(Array.isArray(parsed.places) ? parsed.places : []),
    };
  } catch (error) {
    return defaultStore_();
  }
}

function writeStore_(data) {
  const normalized = {
    members: normalizeNames_(data.members || []),
    expenses: Array.isArray(data.expenses) ? data.expenses : [],
    places: normalizePlaceOrder_(Array.isArray(data.places) ? data.places : []),
  };
  PropertiesService.getScriptProperties().setProperty(STORE_KEY, JSON.stringify(normalized));
}

function defaultStore_() {
  return {
    members: [],
    expenses: [],
    places: [],
  };
}

function normalizeNames_(value) {
  if (Array.isArray(value)) {
    return uniqueNames_(value);
  }

  return uniqueNames_(String(value || '').split(/[\n,、]/));
}

function uniqueNames_(names) {
  const seen = {};
  return names
    .map((name) => String(name || '').trim())
    .filter(Boolean)
    .filter((name) => {
      if (seen[name]) {
        return false;
      }
      seen[name] = true;
      return true;
    });
}

function normalizePlaceOrder_(places) {
  return places
    .slice()
    .sort((a, b) => Number(a.order || 0) - Number(b.order || 0))
    .map((place, index) => Object.assign({}, place, { order: index }));
}
