/** Un solo inicio de sesión: delegación al sistema académico central. */
function mdjAsisLogin(usuario, password) { return mdjLogin(usuario, password); }
function mdjAsisGetSession(token) { return mdjGetSession(token); }
function mdjAsisLogout(token) { return mdjLogout(token); }
function mdjAsisRequireSession_(token) { return mdjRequireSession_(token); }
function mdjAsisRequireRole_(token, roles) { return mdjRequireRole_(token, roles); }
function mdjAsisFindUser_(usuario) { return mdjFindUser_(usuario); }
function mdjAsisValidatePassword_(input, stored) { return mdjValidatePassword_(input, stored); }
function mdjAsisHashSha256_(texto) { return mdjHashTextoSha256_(texto); }