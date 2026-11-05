/* ---------- FACOL: acceso a datos ----------
   Misma interfaz para Supabase (producción) y para el "modo demostración" (datos solo en esta PC,
   para probar la app sin conexión). Las pantallas solo usan window.db. */
(function () {
  const CONN_KEY = 'facol-conexion';
  // Conexión de fábrica: el proyecto de Supabase de FACOL (la clave pública es segura de incluir).
  const DEFAULT_CONN = {mode: 'supabase', url: 'https://tklartbfpakuizilqjow.supabase.co', key: 'sb_publishable_Yj38zzMCgfUi3WmpF0g3Tw_YCNjGinw'};
  const conn = (() => { try { return JSON.parse(localStorage.getItem(CONN_KEY) || 'null') || DEFAULT_CONN; } catch { return DEFAULT_CONN; } })();
  const uid = () => (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2));

  function applyFilters(rows, q = {}) {
    let r = rows;
    for (const [k, v] of Object.entries(q.eq || {})) r = r.filter(x => (x[k] ?? null) === v);
    for (const [k, v] of Object.entries(q.neq || {})) r = r.filter(x => x[k] !== v);
    for (const [k, v] of Object.entries(q.in || {})) r = r.filter(x => v.includes(x[k]));
    for (const [k, v] of Object.entries(q.gte || {})) r = r.filter(x => x[k] >= v);
    for (const [k, v] of Object.entries(q.lte || {})) r = r.filter(x => x[k] <= v);
    if (q.order) { const d = q.desc ? -1 : 1; r = [...r].sort((a, b) => (a[q.order] ?? '') > (b[q.order] ?? '') ? d : (a[q.order] ?? '') < (b[q.order] ?? '') ? -d : 0); }
    return r;
  }

  /* ===================== Supabase ===================== */
  function supa(url, key) {
    const sb = supabase.createClient(url, key, {auth: {storageKey: 'facol-auth', persistSession: true, autoRefreshToken: true, detectSessionInUrl: false}});
    const chk = ({data, error}) => { if (error) throw error; return data; };
    const build = (table, q = {}) => {
      let b = sb.from(table).select(q.select || '*');
      for (const [k, v] of Object.entries(q.eq || {})) b = v === null ? b.is(k, null) : b.eq(k, v);
      for (const [k, v] of Object.entries(q.neq || {})) b = b.neq(k, v);
      for (const [k, v] of Object.entries(q.in || {})) b = b.in(k, v);
      for (const [k, v] of Object.entries(q.gte || {})) b = b.gte(k, v);
      for (const [k, v] of Object.entries(q.lte || {})) b = b.lte(k, v);
      if (q.order) b = b.order(q.order, {ascending: !q.desc});
      return b;
    };
    return {
      mode: 'supabase', sb,
      async session() { const {data} = await sb.auth.getSession(); return data.session ? {email: data.session.user.email} : null; },
      async signIn(email, password) { chk(await sb.auth.signInWithPassword({email, password})); },
      async signOut() { await sb.auth.signOut(); },
      async resetPassword(email) { chk(await sb.auth.resetPasswordForEmail(email)); },
      async isStaff() { return chk(await sb.rpc('is_staff')); },
      async rol() { const {data, error} = await sb.rpc('mi_rol'); return error ? 'admin' : (data || 'operador'); },   // sin la migración de roles, todos son admin
      async list(table, q = {}) {
        if (q.in && Object.values(q.in).some(v => !v.length)) return [];
        const out = []; const page = 1000;
        for (let from = 0; ; from += page) {
          const rows = chk(await build(table, q).range(from, from + page - 1));
          out.push(...rows); if (rows.length < page) break;
        }
        return out;
      },
      async get(table, id) { return chk(await sb.from(table).select('*').eq('id', id).maybeSingle()); },
      async insert(table, rows) { const many = Array.isArray(rows); const d = chk(await sb.from(table).insert(many ? rows : [rows]).select()); return many ? d : d[0]; },
      async update(table, id, patch) { return chk(await sb.from(table).update(patch).eq('id', id).select().maybeSingle()); },
      async updateWhere(table, eq, patch) { let b = sb.from(table).update(patch); for (const [k, v] of Object.entries(eq)) b = Array.isArray(v) ? b.in(k, v) : b.eq(k, v); return chk(await b.select()); },
      async upsert(table, rows, onConflict) { return chk(await sb.from(table).upsert(rows, onConflict ? {onConflict} : undefined).select()); },
      async remove(table, id) { chk(await sb.from(table).delete().eq('id', id)); },
      async removeWhere(table, eq) { let b = sb.from(table).delete(); for (const [k, v] of Object.entries(eq)) b = b.eq(k, v); chk(await b); },
      async rpc(name, args) { return chk(await sb.rpc(name, args)); },
      async upload(path, file) { chk(await sb.storage.from('archivos').upload(path, file, {upsert: true, contentType: file.type || undefined})); return path; },
      async download(path) { return chk(await sb.storage.from('archivos').download(path)); },
      async removeFile(path) { chk(await sb.storage.from('archivos').remove([path])); },
      async exists(path) { const dir = path.split('/').slice(0, -1).join('/'), name = path.split('/').pop(); const d = chk(await sb.storage.from('archivos').list(dir, {search: name})); return d.some(f => f.name === name); },
    };
  }

  /* ===================== Demostración (localStorage) ===================== */
  function demo() {
    const KEY = 'facol-demo-db'; const files = new Map();
    let data = (() => { try { return JSON.parse(localStorage.getItem(KEY) || 'null'); } catch { return null; } })() || {config: [
      {key: 'folio_siguiente', value: 2355}, {key: 'folio_prefijo', value: 'FACOL'}]};
    const save = () => { try { localStorage.setItem(KEY, JSON.stringify(data)); } catch (e) { console.warn(e); } };
    const T = t => (data[t] = data[t] || []);
    const now = () => new Date().toISOString();
    const DEFAULTS = {
      operaciones: {estado: 'abierta', factura_estado: 'pendiente', retorno_entregado: false, requiere_factura: true, modalidad: 'retorno', created_by: 'demo@facol'},
      dispersiones: {estado: 'pendiente', estrategia: 'SCU'}, invitaciones: {estado: 'enviada'}, cambios_cuenta: {estado: 'pendiente'},
      beneficiarios: {activo: true, datos: {}}, clientes: {activo: true, pct_comision: 11.5, requiere_factura: true}, comisionistas: {activo: true}, empresas: {activo: true},
    };
    const withDefaults = (t, r) => ({id: uid(), created_at: now(), ...(DEFAULTS[t] || {}), ...(t === 'invitaciones' ? {token: uid().replace(/-/g, '').slice(0, 24)} : {}), ...r});
    return {
      mode: 'demo',
      async session() { return sessionStorage.getItem('facol-demo-user') ? {email: sessionStorage.getItem('facol-demo-user')} : null; },
      async signIn(email) { sessionStorage.setItem('facol-demo-user', email || 'demo@facol'); },
      async signOut() { sessionStorage.removeItem('facol-demo-user'); },
      async resetPassword() {},
      async isStaff() { return true; },
      async rol() { return sessionStorage.getItem('facol-demo-rol') || 'admin'; },
      async list(table, q) { return applyFilters(T(table), q).map(r => ({...r})); },
      async get(table, id) { const r = T(table).find(x => x.id === id); return r ? {...r} : null; },
      async insert(table, rows) { const many = Array.isArray(rows); const out = (many ? rows : [rows]).map(r => withDefaults(table, r)); T(table).push(...out); save(); return many ? out : out[0]; },
      async update(table, id, patch) { const r = T(table).find(x => x.id === id); if (!r) return null; Object.assign(r, patch); save(); return {...r}; },
      async updateWhere(table, eq, patch) { const rows = T(table).filter(x => Object.entries(eq).every(([k, v]) => Array.isArray(v) ? v.includes(x[k]) : x[k] === v)); rows.forEach(r => Object.assign(r, patch)); save(); return rows; },
      async upsert(table, rows, key = 'id') { const out = []; for (const r of rows) { const e = T(table).find(x => x[key] === r[key]); if (e) { Object.assign(e, r); out.push(e); } else { const n = withDefaults(table, r); T(table).push(n); out.push(n); } } save(); return out; },
      async remove(table, id) { data[table] = T(table).filter(x => x.id !== id); save(); },
      async removeWhere(table, eq) { data[table] = T(table).filter(x => !Object.entries(eq).every(([k, v]) => x[k] === v)); save(); },
      async rpc(name, args) {
        if (name === 'siguiente_folio') {
          const c = T('config').find(x => x.key === 'folio_siguiente'); let n = +c.value;
          while (T('operaciones').some(o => o.numero === n)) n++;
          c.value = n + 1; save(); const p = (T('config').find(x => x.key === 'folio_prefijo') || {}).value || 'FACOL';
          return {numero: n, folio: p + n};
        }
        if (name === 'is_staff') return true;
        throw new Error('Función no disponible en demostración: ' + name);
      },
      async upload(path, file) { files.set(path, file); return path; },
      async download(path) { if (String(path).startsWith('data:')) return (await fetch(path)).blob(); if (!files.has(path)) throw new Error('En modo demostración los archivos solo duran mientras la app está abierta'); return files.get(path); },
      async removeFile(path) { files.delete(path); },
      async exists(path) { return files.has(path); },
      reset() { localStorage.removeItem(KEY); },
    };
  }

  window.db = conn && conn.mode === 'supabase' && conn.url && conn.key ? supa(conn.url, conn.key) : conn && conn.mode === 'demo' ? demo() : null;
  window.dbConn = {
    get: () => conn,
    set(c) { localStorage.setItem(CONN_KEY, JSON.stringify(c)); location.reload(); },
    clear() { localStorage.removeItem(CONN_KEY); location.reload(); },
  };
})();
