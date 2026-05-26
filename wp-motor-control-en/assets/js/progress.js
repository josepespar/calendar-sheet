/* WP Progress adapter — exposes window.SCORM with same interface as scorm.js */
(function (window) {
  'use strict';

  var LS_KEY     = 'mc_en_progress';
  var _saveTimer = null;

  function _cfg()    { return window.MC_EN_Config || {}; }
  function _restUrl(){ return (_cfg().restUrl || '') + 'progress'; }
  function _nonce()  { return _cfg().nonce || ''; }
  function _logged() { return !!_cfg().isLoggedIn; }

  window.SCORM = {
    initialized: true,

    init: function () {
      if (_logged()) {
        fetch(_restUrl(), { headers: { 'X-WP-Nonce': _nonce() } })
          .then(function (r) { return r.json(); })
          .then(function (json) {
            if (json && json.data) {
              var local = localStorage.getItem(LS_KEY) || '';
              if (json.data.length >= local.length) {
                localStorage.setItem(LS_KEY, json.data);
              }
            }
          })
          .catch(function () {});
      }
      return true;
    },

    getValue: function () { return ''; },
    setValue: function () { return true; },
    commit:   function () { return true; },

    finish: function (score) {
      console.info('[Motor Control EN] Course completed. Score:', score);
    },

    saveSuspendData: function (data) {
      var json;
      try { json = JSON.stringify(data); } catch (e) { return; }
      localStorage.setItem(LS_KEY, json);

      if (_logged()) {
        clearTimeout(_saveTimer);
        _saveTimer = setTimeout(function () {
          fetch(_restUrl(), {
            method:  'POST',
            headers: {
              'Content-Type': 'application/json',
              'X-WP-Nonce':   _nonce()
            },
            body: JSON.stringify({ progress_data: json })
          }).catch(function () {});
        }, 800);
      }
    },

    loadSuspendData: function () {
      var json = localStorage.getItem(LS_KEY);
      if (!json) return null;
      try { return JSON.parse(json); } catch (e) { return null; }
    }
  };

})(window);
