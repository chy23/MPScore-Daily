/**
 * =========================================================================
 * LineHubClient.gs - School LINE Hub 客戶端 SDK (Google Apps Script 版)
 * =========================================================================
 */

(function(root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.LineHubClient = factory();
  }
}(typeof self !== 'undefined' ? self : this, function() {

  function LineHubClient(config) {
    if (!config || !config.hubUrl || !config.appId || !config.appSecret) {
      throw new Error("LineHubClient 初始化失敗：必須提供 hubUrl, appId 與 appSecret！");
    }
    this.hubUrl = config.hubUrl.trim();
    this.appId = config.appId.trim();
    this.appSecret = config.appSecret.trim();
    this.timeout = config.timeout || 30000;
  }

  /**
   * 內部 HTTP POST 封裝
   */
  LineHubClient.prototype._post = function(action, extraPayload) {
    var payload = Object.assign({
      action: action,
      appId: this.appId,
      appSecret: this.appSecret
    }, extraPayload || {});

    var jsonStr = JSON.stringify(payload);

    if (typeof UrlFetchApp !== 'undefined') {
      try {
        var resp = UrlFetchApp.fetch(this.hubUrl, {
          method: 'post',
          contentType: 'application/json',
          payload: jsonStr,
          muteHttpExceptions: true
        });
        var code = resp.getResponseCode();
        var body = resp.getContentText();
        var parsed = {};
        try {
          parsed = JSON.parse(body);
        } catch (e) {
          parsed = { status: 'error', code: code, raw: body };
        }
        return parsed;
      } catch (gasErr) {
        return { status: 'error', message: gasErr.toString() };
      }
    }

    if (typeof fetch !== 'undefined') {
      return fetch(this.hubUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: jsonStr
      }).then(function(res) {
        return res.json();
      }).catch(function(err) {
        return { status: 'error', message: err.toString() };
      });
    }

    throw new Error("找不到可用的 HTTP 客戶端 (UrlFetchApp 或 fetch)！");
  };

  /**
   * 1. 測試中樞連線與驗證金鑰
   */
  LineHubClient.prototype.testConnection = function() {
    return this._post('verify_client');
  };

  /**
   * 2. 查詢特定學生的家長綁定狀況
   */
  LineHubClient.prototype.getStudentBinding = function(target) {
    return this._post('get_student_binding', { target: target });
  };

  /**
   * 3. 批次取得全班綁定對照表
   */
  LineHubClient.prototype.getAllBindingsSummary = function() {
    return this._post('get_all_bindings_summary');
  };

  /**
   * 4. 同步學生清冊名單到中樞
   */
  LineHubClient.prototype.registerStudents = function(studentsList) {
    return this._post('register_students', { students: studentsList });
  };

  /**
   * 5. 發送單一推播通知
   */
  LineHubClient.prototype.sendPushNotification = function(options) {
    return this._post('send_push', {
      target: {
        seatNo: options.seatNo,
        studentId: options.studentId,
        studentName: options.studentName,
        token: options.token,
        title: options.title,
        body: options.body,
        url: options.url,
        customMessage: options.customMessage
      }
    });
  };

  /**
   * 6. 批次發送多人推播
   */
  LineHubClient.prototype.sendBatchPush = function(targets, messageTemplate) {
    return this._post('send_push', {
      targets: targets,
      messageTemplate: messageTemplate
    });
  };

  return LineHubClient;
}));
