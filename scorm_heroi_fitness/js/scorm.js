/* SCORM 1.2 API Wrapper */
(function (window) {
  'use strict';

  var SCORM = {
    API: null,
    initialized: false,

    _findAPI: function (win) {
      var tries = 0;
      while (!win.API && win.parent && win.parent !== win) {
        tries++;
        if (tries > 7) break;
        win = win.parent;
      }
      return win.API || null;
    },

    _getAPI: function () {
      var api = this._findAPI(window);
      if (!api && window.opener) {
        api = this._findAPI(window.opener);
      }
      return api;
    },

    init: function () {
      this.API = this._getAPI();
      if (!this.API) {
        console.info('[SCORM] API no trobada. Mode standalone actiu.');
        return false;
      }
      var ok = this.API.LMSInitialize('');
      if (ok === 'true' || ok === true) {
        this.initialized = true;
        this.setValue('cmi.core.lesson_status', 'incomplete');
        this.commit();
        console.info('[SCORM] LMS inicialitzat correctament.');
        return true;
      }
      console.warn('[SCORM] LMSInitialize ha fallat.');
      return false;
    },

    getValue: function (key) {
      if (!this.API || !this.initialized) return '';
      return this.API.LMSGetValue(key) || '';
    },

    setValue: function (key, value) {
      if (!this.API || !this.initialized) return false;
      var r = this.API.LMSSetValue(key, String(value));
      return r === 'true' || r === true;
    },

    commit: function () {
      if (!this.API || !this.initialized) return false;
      var r = this.API.LMSCommit('');
      return r === 'true' || r === true;
    },

    finish: function (score) {
      if (!this.API || !this.initialized) return false;
      this.setValue('cmi.core.score.raw', String(Math.round(score)));
      this.setValue('cmi.core.score.min', '0');
      this.setValue('cmi.core.score.max', '100');
      this.setValue('cmi.core.lesson_status', 'completed');
      this.commit();
      this.API.LMSFinish('');
      this.initialized = false;
      console.info('[SCORM] LMS finalitzat. Puntuació:', score);
      return true;
    },

    saveSuspendData: function (data) {
      try {
        var json = JSON.stringify(data);
        this.setValue('cmi.suspend_data', json);
        this.commit();
      } catch (e) {
        console.warn('[SCORM] Error guardant suspend_data:', e);
      }
    },

    loadSuspendData: function () {
      var json = this.getValue('cmi.suspend_data');
      if (!json) return null;
      try {
        return JSON.parse(json);
      } catch (e) {
        return null;
      }
    }
  };

  window.SCORM = SCORM;
})(window);
