// 取得違規法條清單
function fetchLawList() {
  return new Promise((resolve, reject) => {
    $.ajax({
      url: "api/common/laws.php",
      type: "GET",
      dataType: "json",
      success: function (res) {
        if (res.returnCode == 200) {
          resolve(res.data);
        } else {
          reject(new Error(res.message || "取得違規法條失敗"));
        }
      },
      error: function (xhr) {
        reject(new Error(xhr?.responseJSON?.message || xhr.responseText));
      },
    });
  });
}

// 取得車種清單
function fetchVehicleTypeList() {
  return new Promise((resolve, reject) => {
    $.ajax({
      url: "api/common/vehicle_types.php",
      type: "GET",
      dataType: "json",
      success: function (res) {
        if (res.returnCode == 200) {
          resolve(res.data);
        } else {
          reject(new Error(res.message || "取得車種清單失敗"));
        }
      },
      error: function (xhr) {
        reject(new Error(xhr?.responseJSON?.message || xhr.responseText));
      },
    });
  });
}

// 取得輔助車種清單
function fetchVehicleTypeAideList() {
  return new Promise((resolve, reject) => {
    $.ajax({
      url: "api/common/vehicle_types_aide.php",
      type: "GET",
      dataType: "json",
      success: function (res) {
        if (res.returnCode == 200) {
          resolve(res.data);
        } else {
          reject(new Error(res.message || "取得車種清單失敗"));
        }
      },
      error: function (xhr) {
        reject(new Error(xhr?.responseJSON?.message || xhr.responseText));
      },
    });
  });
}

// 取得簽收狀況清單
function fetchSigningStatusList() {
  return new Promise((resolve, reject) => {
    $.ajax({
      url: "api/common/signing_status.php",
      type: "GET",
      dataType: "json",
      success: function (res) {
        if (res.returnCode == 200) {
          resolve(res.data);
        } else {
          reject(new Error(res.message || "取得簽收狀況清單失敗"));
        }
      },
      error: function (xhr) {
        reject(new Error(xhr?.responseJSON?.message || xhr.responseText));
      },
    });
  });
}

// 取得代保管物清單
function fetchDepositoryList() {
  return new Promise((resolve, reject) => {
    $.ajax({
      url: "api/common/depository.php",
      type: "GET",
      dataType: "json",
      success: function (res) {
        if (res.returnCode == 200) {
          resolve(res.data);
        } else {
          reject(new Error(res.message || "取得代保管物清單失敗"));
        }
      },
      error: function (xhr) {
        reject(new Error(xhr?.responseJSON?.message || xhr.responseText));
      },
    });
  });
}

// 取得保險證清單
function fetchInsuranceCardList() {
  return new Promise((resolve, reject) => {
    $.ajax({
      url: "api/common/insurance_card.php",
      type: "GET",
      dataType: "json",
      success: function (res) {
        if (res.returnCode == 200) {
          resolve(res.data);
        } else {
          reject(new Error(res.message || "取得保險證清單失敗"));
        }
      },
      error: function (xhr) {
        reject(new Error(xhr?.responseJSON?.message || xhr.responseText));
      },
    });
  });
}

// 取得微電車種類清單
function fetchElectricScooterTypeList() {
  return new Promise((resolve, reject) => {
    $.ajax({
      url: "api/common/electric_scooter_types.php",
      type: "GET",
      dataType: "json",
      success: function (res) {
        if (res.returnCode == 200) {
          resolve(res.data);
        } else {
          reject(new Error(res.message || "取得微電車種類清單失敗"));
        }
      },
      error: function (xhr) {
        reject(new Error(xhr?.responseJSON?.message || xhr.responseText));
      },
    });
  });
}
