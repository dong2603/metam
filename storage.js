const STORAGE_KEY = 'leaveCheckerStoreExcel_v1';

const LEAVE_TYPE_RULES = {
  '일반 연차': { label: '일반 연차', daysUsed: 1 },
  '반일 연차': { label: '반일 연차', daysUsed: 0.5 },
  '생일 연차': { label: '생일 연차', daysUsed: 0 },
  '공가 연차': { label: '공가 연차', daysUsed: 0 }
};

const LEAVE_TYPE_ALIASES = {
  '연차': '일반 연차',
  '일반연차': '일반 연차',
  '반차': '반일 연차',
  '반일연차': '반일 연차',
  '생일연차': '생일 연차',
  '공가': '공가 연차',
  '공가연차': '공가 연차'
};

const defaultStore = {
  employees: [
    {
      id: 'emp_1',
      name: '홍길동',
      rrnFront: '900101',
      department: '개발팀',
      hireDate: '2021-03-02',
      annualLeaveTotal: 15
    }
  ],
  histories: [
    {
      id: 'his_1',
      employeeId: 'emp_1',
      leaveDate: '2026-01-20',
      leaveType: '일반 연차',
      daysUsed: 1,
      reason: '개인 일정'
    },
    {
      id: 'his_2',
      employeeId: 'emp_1',
      leaveDate: '2026-03-14',
      leaveType: '반일 연차',
      daysUsed: 0.5,
      reason: '병원 방문'
    },
    {
      id: 'his_3',
      employeeId: 'emp_1',
      leaveDate: '2026-05-15',
      leaveType: '생일 연차',
      daysUsed: 0,
      reason: '생일 휴가'
    }
  ]
};

function clone(data) {
  return JSON.parse(JSON.stringify(data));
}

function uid(prefix) {
  return prefix + '_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8);
}

function normalizeDate(value) {
  if (!value) return '';
  if (typeof value === 'number') {
    if (window.XLSX && XLSX.SSF) {
      const date = XLSX.SSF.parse_date_code(value);
      if (date) {
        return `${date.y}-${String(date.m).padStart(2, '0')}-${String(date.d).padStart(2, '0')}`;
      }
    }
    return String(value);
  }
  const text = String(value).trim();
  if (!text) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  if (/^\d{4}\.\d{2}\.\d{2}$/.test(text)) return text.replace(/\./g, '-');
  if (/^\d{4}\/\d{2}\/\d{2}$/.test(text)) return text.replace(/\//g, '-');
  return text;
}

function normalizeLeaveType(value) {
  const text = String(value || '').trim();
  if (LEAVE_TYPE_RULES[text]) return text;
  if (LEAVE_TYPE_ALIASES[text] && LEAVE_TYPE_RULES[LEAVE_TYPE_ALIASES[text]]) {
    return LEAVE_TYPE_ALIASES[text];
  }
  return '일반 연차';
}

function getDeductionDaysByType(leaveType) {
  const normalized = normalizeLeaveType(leaveType);
  return Number(LEAVE_TYPE_RULES[normalized].daysUsed);
}

function normalizeEmployeeRecord(employee) {
  return {
    id: employee.id || uid('emp'),
    name: String(employee.name || '').trim(),
    rrnFront: String(employee.rrnFront || '').trim(),
    department: String(employee.department || '').trim(),
    hireDate: normalizeDate(employee.hireDate || ''),
    annualLeaveTotal: Number(employee.annualLeaveTotal || 0)
  };
}

function normalizeHistoryRecord(history) {
  const leaveType = normalizeLeaveType(history.leaveType);
  return {
    id: history.id || uid('his'),
    employeeId: String(history.employeeId || '').trim(),
    leaveDate: normalizeDate(history.leaveDate || ''),
    leaveType,
    daysUsed: getDeductionDaysByType(leaveType),
    reason: String(history.reason || '').trim()
  };
}

function seedStoreIfNeeded() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) localStorage.setItem(STORAGE_KEY, JSON.stringify(defaultStore));
}

function loadStore() {
  seedStoreIfNeeded();
  try {
    const parsed = JSON.parse(localStorage.getItem(STORAGE_KEY));
    if (!parsed || !Array.isArray(parsed.employees) || !Array.isArray(parsed.histories)) {
      throw new Error('invalid');
    }

    const normalizedStore = {
      employees: parsed.employees.map(normalizeEmployeeRecord).filter((item) => item.name && /^\d{6}$/.test(item.rrnFront)),
      histories: parsed.histories.map(normalizeHistoryRecord).filter((item) => item.employeeId && item.leaveDate)
    };

    localStorage.setItem(STORAGE_KEY, JSON.stringify(normalizedStore));
    return normalizedStore;
  } catch (error) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(defaultStore));
    return clone(defaultStore);
  }
}

function saveStore(store) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(store));
}

function resetStore() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(defaultStore));
  return loadStore();
}

function setStore(store) {
  if (!store || !Array.isArray(store.employees) || !Array.isArray(store.histories)) {
    throw new Error('형식이 올바르지 않습니다. employees 와 histories 배열이 필요합니다.');
  }
  const normalizedStore = {
    employees: store.employees.map(normalizeEmployeeRecord).filter((item) => item.name && /^\d{6}$/.test(item.rrnFront)),
    histories: store.histories.map(normalizeHistoryRecord).filter((item) => item.employeeId && item.leaveDate)
  };
  saveStore(normalizedStore);
  return loadStore();
}

function calculateUsedLeave(store, employeeId) {
  return store.histories
    .filter((item) => item.employeeId === employeeId)
    .reduce((sum, item) => sum + getDeductionDaysByType(item.leaveType), 0);
}

function getEmployeeSummary(store, employee) {
  const total = Number(employee.annualLeaveTotal || 0);
  const used = calculateUsedLeave(store, employee.id);
  return {
    ...employee,
    annualLeaveTotal: total,
    annualLeaveUsed: used,
    annualLeaveRemaining: total - used
  };
}

function getEmployees(store = loadStore()) {
  return [...store.employees].sort((a, b) => a.name.localeCompare(b.name, 'ko'));
}

function getHistoriesByEmployeeId(employeeId, store = loadStore()) {
  return store.histories
    .filter((item) => item.employeeId === employeeId)
    .map(normalizeHistoryRecord)
    .sort((a, b) => String(b.leaveDate).localeCompare(String(a.leaveDate)));
}

function findEmployee(name, rrnFront, store = loadStore()) {
  return store.employees.find((item) => item.name.trim() === name.trim() && item.rrnFront === rrnFront);
}

function upsertEmployee(payload) {
  const store = loadStore();
  const name = String(payload.name || '').trim();
  const rrnFront = String(payload.rrnFront || '').trim();
  const department = String(payload.department || '').trim();
  const hireDate = normalizeDate(payload.hireDate);
  const annualLeaveTotal = Number(payload.annualLeaveTotal || 0);

  const existing = store.employees.find((item) => item.name === name && item.rrnFront === rrnFront);
  if (existing) {
    existing.department = department;
    existing.hireDate = hireDate;
    existing.annualLeaveTotal = annualLeaveTotal;
  } else {
    store.employees.push({
      id: uid('emp'),
      name,
      rrnFront,
      department,
      hireDate,
      annualLeaveTotal
    });
  }
  saveStore(store);
  return loadStore();
}

function addHistory(payload) {
  const store = loadStore();
  store.histories.push(normalizeHistoryRecord({
    id: uid('his'),
    employeeId: payload.employeeId,
    leaveDate: payload.leaveDate,
    leaveType: payload.leaveType,
    reason: payload.reason
  }));
  saveStore(store);
  return loadStore();
}

function removeEmployee(employeeId) {
  const store = loadStore();
  store.employees = store.employees.filter((item) => item.id !== employeeId);
  store.histories = store.histories.filter((item) => item.employeeId !== employeeId);
  saveStore(store);
  return loadStore();
}

function removeHistory(historyId) {
  const store = loadStore();
  store.histories = store.histories.filter((item) => item.id !== historyId);
  saveStore(store);
  return loadStore();
}

function exportStore() {
  return JSON.stringify(loadStore(), null, 2);
}

function importStore(jsonText) {
  const parsed = JSON.parse(jsonText);
  return setStore(parsed);
}

function normalizeImportedWorkbook(workbook, replaceAll) {
  const sheetNames = workbook.SheetNames || [];
  const employeeSheetName = sheetNames.find((name) => name.toLowerCase() === 'employees') || sheetNames[0];
  const historySheetName = sheetNames.find((name) => name.toLowerCase() === 'histories') || sheetNames[1];

  if (!employeeSheetName) {
    throw new Error('employees 시트가 필요합니다.');
  }

  const employeeRows = XLSX.utils.sheet_to_json(workbook.Sheets[employeeSheetName], { defval: '' });
  const historyRows = historySheetName ? XLSX.utils.sheet_to_json(workbook.Sheets[historySheetName], { defval: '' }) : [];

  const baseStore = replaceAll ? { employees: [], histories: [] } : loadStore();
  const employeeMap = new Map(baseStore.employees.map((item) => [`${item.name}__${item.rrnFront}`, item]));

  employeeRows.forEach((row) => {
    const name = String(row.name || row.Name || '').trim();
    const rrnFront = String(row.rrnFront || row.rrn_front || row.birth6 || '').trim();
    if (!name || !/^\d{6}$/.test(rrnFront)) return;

    const key = `${name}__${rrnFront}`;
    const current = employeeMap.get(key);
    employeeMap.set(key, {
      id: current ? current.id : uid('emp'),
      name,
      rrnFront,
      department: String(row.department || row.Dept || '').trim(),
      hireDate: normalizeDate(row.hireDate || row.hire_date || ''),
      annualLeaveTotal: Number(row.annualLeaveTotal || row.annual_leave_total || 0)
    });
  });

  const employees = Array.from(employeeMap.values()).map(normalizeEmployeeRecord);
  const employeesByKey = new Map(employees.map((item) => [`${item.name}__${item.rrnFront}`, item]));
  const histories = replaceAll ? [] : [...baseStore.histories.map(normalizeHistoryRecord)];

  historyRows.forEach((row) => {
    const name = String(row.name || row.Name || '').trim();
    const rrnFront = String(row.rrnFront || row.rrn_front || row.birth6 || '').trim();
    const employee = employeesByKey.get(`${name}__${rrnFront}`);
    if (!employee) return;

    const leaveDate = normalizeDate(row.leaveDate || row.leave_date || '');
    if (!leaveDate) return;

    histories.push(normalizeHistoryRecord({
      id: uid('his'),
      employeeId: employee.id,
      leaveDate,
      leaveType: row.leaveType || row.leave_type || '일반 연차',
      reason: String(row.reason || '').trim()
    }));
  });

  return { employees, histories };
}

window.LeaveStore = {
  LEAVE_TYPE_RULES,
  seedStoreIfNeeded,
  loadStore,
  saveStore,
  setStore,
  resetStore,
  getEmployees,
  getEmployeeSummary,
  getHistoriesByEmployeeId,
  findEmployee,
  upsertEmployee,
  addHistory,
  removeEmployee,
  removeHistory,
  exportStore,
  importStore,
  normalizeImportedWorkbook,
  normalizeLeaveType,
  getDeductionDaysByType
};
