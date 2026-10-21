/**
 * Autenticación y sesiones
 */

function mdjLogin(usuario, password) {
  usuario = String(usuario || '').trim();
  password = String(password || '');

  if (!usuario || !password) {
    return { ok: false, message: 'Debes ingresar usuario y contraseña.' };
  }

  const user = mdjFindUser_(usuario);
  if (!user) {
    return { ok: false, message: 'Usuario no encontrado.' };
  }
  if (!user.activo) {
    return { ok: false, message: 'El usuario existe pero está inactivo.' };
  }
  if (!mdjValidatePassword_(password, user.password)) {
    return { ok: false, message: 'Contraseña incorrecta.' };
  }

  const token = Utilities.getUuid();
  const session = {
    token: token,
    createdAt: new Date().toISOString(),
    user: {
      usuario: user.usuario,
      nombre: user.nombre,
      rol: user.rol,
      activo: user.activo
    }
  };

  CacheService.getScriptCache().put('MDJ_SESSION_' + token, JSON.stringify(session), MDJ_APP.sessionTtlSeconds);

  return {
    ok: true,
    message: 'Inicio de sesión correcto.',
    token: token,
    user: session.user,
    redirectPage: user.rol === 'admin' ? 'admin' : 'maestro'
  };
}

function mdjGetSession(token) {
  try {
    const session = mdjRequireSession_(token);
    return { ok: true, user: session.user };
  } catch (error) {
    return { ok: false, message: error.message };
  }
}

function mdjLogout(token) {
  token = String(token || '').trim();
  if (token) {
    CacheService.getScriptCache().remove('MDJ_SESSION_' + token);
  }
  return { ok: true };
}

function mdjHashTextoSha256(texto) {
  return 'sha256:' + mdjHashTextoSha256_(texto);
}

function mdjRequireSession_(token) {
  token = String(token || '').trim();
  if (!token) {
    throw new Error('Sesión no enviada.');
  }
  const raw = CacheService.getScriptCache().get('MDJ_SESSION_' + token);
  if (!raw) {
    throw new Error('La sesión expiró o no existe.');
  }
  const session = JSON.parse(raw);
  if (!session || !session.user || !session.user.usuario) {
    throw new Error('La sesión es inválida.');
  }
  return session;
}

function mdjRequireRole_(token, allowedRoles) {
  const session = mdjRequireSession_(token);
  const role = String(session.user.rol || '').toLowerCase();
  const normalized = (allowedRoles || []).map(r => String(r).toLowerCase());
  if (normalized.indexOf(role) === -1) {
    throw new Error('No tienes permisos para realizar esta acción.');
  }
  return session;
}

function mdjFindUser_(usuarioBuscado) {
  const sh = mdjGetControlSheet_(MDJ_APP.hojaUsuarios);
  const values = sh.getDataRange().getDisplayValues();
  if (!values || values.length < 2) return null;

  const header = mdjBuildHeaderMap_(values[0]);
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const usuario = String(row[header.usuario] || '').trim();
    if (!usuario) continue;
    if (mdjNormalizeText_(usuario) !== mdjNormalizeText_(usuarioBuscado)) continue;

    const activoCell = String(row[header.activo] || 'SI').trim();
    return {
      usuario: usuario,
      password: String(row[header.password] || ''),
      nombre: String(row[header.nombre] || usuario),
      rol: mdjNormalizeRole_(row[header.rol]),
      activo: mdjIsTruthy_(activoCell)
    };
  }
  return null;
}

function mdjValidatePassword_(passwordInput, storedPassword) {
  const stored = String(storedPassword || '').trim();
  if (!stored) return false;

  if (/^sha256:/i.test(stored)) {
    return mdjHashTextoSha256_(passwordInput) === stored.replace(/^sha256:/i, '');
  }
  return stored === passwordInput;
}

function mdjHashTextoSha256_(texto) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, texto, Utilities.Charset.UTF_8);
  return bytes.map(function(b) {
    const v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}