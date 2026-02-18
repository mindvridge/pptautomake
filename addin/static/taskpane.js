/**
 * PPT AutoMake - Task Pane 클라이언트
 *
 * Office.js와 Flask 백엔드 API를 연결하여
 * PowerPoint 내 Task Pane에서 도식 분석/변환을 수행한다.
 *
 * 두 가지 모드 지원:
 * 1) Office.js 연동: 현재 열린 PPT에서 직접 파일을 가져오고 결과를 삽입
 * 2) 파일 업로드: 웹 브라우저에서 PPTX 파일을 직접 업로드
 */

(function () {
  "use strict";

  // --- 상태 ---
  var selectedFile = null; // File 객체 (업로드 모드)
  var currentFileBlob = null; // Blob (Office.js 모드)
  var isOfficeReady = false;
  var API_BASE = window.location.origin;

  // --- DOM 요소 ---
  var els = {};
  function initElements() {
    els.connectionStatus = document.getElementById("connection-status");
    els.statusText = els.connectionStatus.querySelector(".status-text");
    els.currentFileSection = document.getElementById("current-file-section");
    els.btnUseCurrent = document.getElementById("btn-use-current");
    els.currentFileInfo = document.getElementById("current-file-info");
    els.currentFileName = document.getElementById("current-file-name");
    els.btnClearCurrent = document.getElementById("btn-clear-current");
    els.uploadArea = document.getElementById("upload-area");
    els.uploadContent = document.getElementById("upload-content");
    els.fileInput = document.getElementById("file-input");
    els.fileInfo = document.getElementById("file-info");
    els.fileName = document.getElementById("file-name");
    els.btnClear = document.getElementById("btn-clear");
    els.slidesInput = document.getElementById("slides-input");
    els.insertOption = document.getElementById("insert-option");
    els.autoInsert = document.getElementById("auto-insert");
    els.btnAnalyze = document.getElementById("btn-analyze");
    els.btnProcess = document.getElementById("btn-process");
    els.progressSection = document.getElementById("progress-section");
    els.progressFill = document.getElementById("progress-fill");
    els.progressText = document.getElementById("progress-text");
    els.resultSection = document.getElementById("result-section");
    els.resultContent = document.getElementById("result-content");
  }

  // --- 초기화 ---
  document.addEventListener("DOMContentLoaded", function () {
    initElements();
    bindEvents();
    checkServerHealth();

    // Office.js 초기화 (PPT 안에서만 동작)
    if (typeof Office !== "undefined") {
      Office.onReady(function (info) {
        if (info.host === Office.HostType.PowerPoint) {
          isOfficeReady = true;
          els.currentFileSection.hidden = false;
          els.insertOption.hidden = false;
        }
      });
    }
  });

  // --- 이벤트 바인딩 ---
  function bindEvents() {
    // 현재 파일 사용 (Office.js)
    els.btnUseCurrent.addEventListener("click", getCurrentFile);
    els.btnClearCurrent.addEventListener("click", clearCurrentFile);

    // 파일 업로드
    els.uploadArea.addEventListener("click", function () {
      els.fileInput.click();
    });
    els.fileInput.addEventListener("change", handleFileSelect);
    els.btnClear.addEventListener("click", clearFile);

    // 드래그 앤 드롭
    els.uploadArea.addEventListener("dragover", function (e) {
      e.preventDefault();
      els.uploadArea.classList.add("dragover");
    });
    els.uploadArea.addEventListener("dragleave", function () {
      els.uploadArea.classList.remove("dragover");
    });
    els.uploadArea.addEventListener("drop", function (e) {
      e.preventDefault();
      els.uploadArea.classList.remove("dragover");
      if (e.dataTransfer.files.length > 0) {
        setFile(e.dataTransfer.files[0]);
      }
    });

    // 실행 버튼
    els.btnAnalyze.addEventListener("click", function () {
      runPipeline("analyze");
    });
    els.btnProcess.addEventListener("click", function () {
      runPipeline("process");
    });
  }

  // ============================================================
  // Office.js 파일 가져오기 (getFileAsync)
  // ============================================================
  function getCurrentFile() {
    if (!isOfficeReady) return;

    els.btnUseCurrent.disabled = true;
    els.btnUseCurrent.textContent = "파일 가져오는 중...";

    Office.context.document.getFileAsync(
      Office.FileType.Compressed,
      { sliceSize: 65536 },
      function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          els.btnUseCurrent.disabled = false;
          els.btnUseCurrent.innerHTML =
            '<span class="btn-icon">📄</span> 현재 열린 프레젠테이션 사용';
          showError("파일을 가져올 수 없습니다: " + result.error.message);
          return;
        }

        var file = result.value;
        var sliceCount = file.sliceCount;
        var slicesReceived = 0;
        var allSlices = [];

        for (var i = 0; i < sliceCount; i++) {
          file.getSliceAsync(i, function (sliceResult) {
            if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
              allSlices[sliceResult.value.index] = sliceResult.value.data;
              slicesReceived++;

              if (slicesReceived === sliceCount) {
                file.closeAsync();

                // 모든 슬라이스를 합쳐서 Blob 생성
                var docData = [];
                for (var j = 0; j < sliceCount; j++) {
                  docData = docData.concat(allSlices[j]);
                }
                var byteArray = new Uint8Array(docData);
                currentFileBlob = new Blob([byteArray], {
                  type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                });

                // UI 업데이트
                selectedFile = null;
                els.fileInput.value = "";
                els.uploadContent.hidden = false;
                els.fileInfo.hidden = true;

                els.btnUseCurrent.hidden = true;
                els.currentFileInfo.hidden = false;
                els.currentFileName.textContent =
                  "현재 프레젠테이션 (" +
                  Math.round(currentFileBlob.size / 1024) +
                  "KB)";
                updateButtons();
              }
            }
          });
        }
      }
    );
  }

  function clearCurrentFile(e) {
    if (e) e.stopPropagation();
    currentFileBlob = null;
    els.btnUseCurrent.hidden = false;
    els.btnUseCurrent.disabled = false;
    els.btnUseCurrent.innerHTML =
      '<span class="btn-icon">📄</span> 현재 열린 프레젠테이션 사용';
    els.currentFileInfo.hidden = true;
    updateButtons();
  }

  // ============================================================
  // 파일 업로드 관리
  // ============================================================
  function handleFileSelect(e) {
    if (e.target.files.length > 0) {
      setFile(e.target.files[0]);
    }
  }

  function setFile(file) {
    if (!file.name.toLowerCase().endsWith(".pptx")) {
      showError("PPTX 파일만 지원합니다.");
      return;
    }
    selectedFile = file;
    currentFileBlob = null;
    clearCurrentFile();

    els.uploadContent.hidden = true;
    els.fileInfo.hidden = false;
    els.fileName.textContent = file.name;
    updateButtons();
  }

  function clearFile(e) {
    if (e) e.stopPropagation();
    selectedFile = null;
    els.fileInput.value = "";
    els.uploadContent.hidden = false;
    els.fileInfo.hidden = true;
    updateButtons();
  }

  function updateButtons() {
    var hasFile = selectedFile !== null || currentFileBlob !== null;
    els.btnAnalyze.disabled = !hasFile;
    els.btnProcess.disabled = !hasFile;
  }

  // --- 서버 연결 ---
  function checkServerHealth() {
    fetch(API_BASE + "/api/health")
      .then(function (res) {
        if (res.ok) {
          setConnectionStatus("connected", "서버 연결됨");
        } else {
          setConnectionStatus("error", "서버 응답 오류");
        }
      })
      .catch(function () {
        setConnectionStatus("error", "서버에 연결할 수 없습니다");
      });
  }

  function setConnectionStatus(state, text) {
    els.connectionStatus.className = "status-bar status-" + state;
    els.statusText.textContent = text;
  }

  // ============================================================
  // 파이프라인 실행
  // ============================================================
  function runPipeline(mode) {
    var fileToSend = selectedFile || currentFileBlob;
    if (!fileToSend) return;

    // UI 상태 변경
    els.btnAnalyze.disabled = true;
    els.btnProcess.disabled = true;
    els.progressSection.hidden = false;
    els.resultSection.hidden = true;

    var stageText = mode === "analyze" ? "분석" : "분석 + 변환";
    setProgress(-1, stageText + " 진행 중...");

    var formData = new FormData();
    var fileName = selectedFile ? selectedFile.name : "current.pptx";
    formData.append("file", fileToSend, fileName);

    var slides = els.slidesInput.value.trim();
    if (slides) {
      formData.append("slides", slides);
    }

    // Office.js 모드에서 자동 삽입 요청 시 base64 응답 요청
    var wantInsert =
      isOfficeReady &&
      currentFileBlob &&
      mode === "process" &&
      els.autoInsert.checked;
    if (wantInsert) {
      formData.append("return_base64", "true");
    }

    var endpoint = mode === "analyze" ? "/api/analyze" : "/api/process";

    fetch(API_BASE + endpoint, {
      method: "POST",
      body: formData,
    })
      .then(function (res) {
        return res.json().then(function (data) {
          return { ok: res.ok, data: data };
        });
      })
      .then(function (result) {
        if (!result.ok) {
          els.progressSection.hidden = true;
          showError(result.data.error || "알 수 없는 오류가 발생했습니다.");
          return;
        }

        if (mode === "analyze") {
          els.progressSection.hidden = true;
          showAnalyzeResult(result.data);
        } else if (wantInsert && result.data.base64) {
          // Office.js로 슬라이드 삽입
          setProgress(-1, "PPT에 슬라이드 삽입 중...");
          insertSlidesFromBase64(result.data.base64, result.data);
        } else {
          els.progressSection.hidden = true;
          showProcessResult(result.data);
        }
      })
      .catch(function (err) {
        els.progressSection.hidden = true;
        showError("네트워크 오류: " + err.message);
      })
      .finally(function () {
        updateButtons();
      });
  }

  // ============================================================
  // Office.js 슬라이드 삽입 (insertSlidesFromBase64)
  // ============================================================
  function insertSlidesFromBase64(base64Data, resultData) {
    PowerPoint.run(function (context) {
      context.presentation.insertSlidesFromBase64(base64Data);
      return context.sync();
    })
      .then(function () {
        els.progressSection.hidden = true;
        showInsertResult(resultData);
      })
      .catch(function (err) {
        els.progressSection.hidden = true;
        showProcessResult(resultData);
        showError("슬라이드 삽입 실패: " + err.message + " (파일 다운로드를 이용하세요)");
      });
  }

  // --- 프로그레스 ---
  function setProgress(percent, text) {
    if (percent < 0) {
      els.progressFill.style.width = "";
      els.progressFill.classList.add("indeterminate");
    } else {
      els.progressFill.classList.remove("indeterminate");
      els.progressFill.style.width = percent + "%";
    }
    els.progressText.textContent = text;
  }

  // ============================================================
  // 결과 표시
  // ============================================================
  function showAnalyzeResult(data) {
    els.resultSection.hidden = false;
    var html = "";

    html +=
      '<div class="result-summary success">' +
      data.total_slides +
      "개 슬라이드 분석 완료 (재구성 대상: " +
      data.rebuild_count +
      "개)</div>";

    data.slides.forEach(function (slide) {
      html += '<div class="slide-card">';
      html += '<div class="slide-card-header">';
      html += '<span class="slide-number">슬라이드 ' + slide.index + "</span>";

      if (slide.needs_rebuild) {
        html += '<span class="badge badge-rebuild">변환 대상</span>';
      } else {
        html += '<span class="badge badge-skip">유지</span>';
      }

      html += "</div>";
      html +=
        '<div class="slide-details">요소: ' +
        slide.element_count +
        "개</div>";

      if (slide.diagrams && slide.diagrams.length > 0) {
        slide.diagrams.forEach(function (d) {
          html +=
            '<span class="diagram-tag">' +
            d.diagram_type +
            " (" +
            Math.round(d.confidence * 100) +
            "%)</span>";
        });
      }

      html += "</div>";
    });

    els.resultContent.innerHTML = html;
  }

  function showProcessResult(data) {
    els.resultSection.hidden = false;
    var html = "";

    if (data.download_url) {
      html +=
        '<div class="result-summary success">' + data.message + "</div>";

      if (data.details) {
        data.details.forEach(function (d) {
          html += '<div class="slide-card">';
          html +=
            '<span class="slide-number">슬라이드 ' + d.slide + "</span> ";
          html +=
            '<span class="diagram-tag">' +
            d.diagram_type +
            " (" +
            d.node_count +
            "개 노드)</span>";
          html += "</div>";
        });
      }

      html +=
        '<a class="btn-download" href="' +
        API_BASE +
        data.download_url +
        '" download>변환된 파일 다운로드</a>';
    } else {
      html +=
        '<div class="result-summary">' +
        (data.message || "재구성할 도식이 없습니다.") +
        "</div>";
    }

    els.resultContent.innerHTML = html;
  }

  function showInsertResult(data) {
    els.resultSection.hidden = false;
    var html = "";

    html += '<div class="result-summary success">변환 완료 - 슬라이드가 삽입되었습니다</div>';

    if (data.details) {
      data.details.forEach(function (d) {
        html += '<div class="slide-card">';
        html += '<span class="slide-number">슬라이드 ' + d.slide + "</span> ";
        html +=
          '<span class="diagram-tag">' +
          d.diagram_type +
          " (" +
          d.node_count +
          "개 노드)</span>";
        html += "</div>";
      });
    }

    html += '<div class="insert-success-note">현재 프레젠테이션 끝에 재구성된 슬라이드가 추가되었습니다.</div>';

    els.resultContent.innerHTML = html;
  }

  function showError(message) {
    els.resultSection.hidden = false;
    els.resultContent.innerHTML =
      '<div class="result-summary error">' + message + "</div>";
  }
})();
