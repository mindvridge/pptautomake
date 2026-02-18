/**
 * PPT AutoMake - Task Pane 클라이언트
 *
 * Office.js와 Flask 백엔드 API를 연결하여
 * PowerPoint 내 Task Pane에서 도식 분석/변환을 수행한다.
 */

(function () {
  "use strict";

  // --- 상태 ---
  let selectedFile = null;
  let isOfficeReady = false;
  const API_BASE = window.location.origin;

  // --- DOM 요소 ---
  const els = {};
  function initElements() {
    els.connectionStatus = document.getElementById("connection-status");
    els.statusText = els.connectionStatus.querySelector(".status-text");
    els.uploadArea = document.getElementById("upload-area");
    els.uploadContent = document.getElementById("upload-content");
    els.fileInput = document.getElementById("file-input");
    els.fileInfo = document.getElementById("file-info");
    els.fileName = document.getElementById("file-name");
    els.btnClear = document.getElementById("btn-clear");
    els.pptNotice = document.getElementById("ppt-notice");
    els.slidesInput = document.getElementById("slides-input");
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
          els.pptNotice.hidden = false;
        }
      });
    }
  });

  // --- 이벤트 바인딩 ---
  function bindEvents() {
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

    // 버튼
    els.btnAnalyze.addEventListener("click", function () {
      runPipeline("analyze");
    });
    els.btnProcess.addEventListener("click", function () {
      runPipeline("process");
    });
  }

  // --- 파일 관리 ---
  function handleFileSelect(e) {
    if (e.target.files.length > 0) {
      setFile(e.target.files[0]);
    }
  }

  function setFile(file) {
    if (!file.name.toLowerCase().endsWith(".pptx")) {
      alert("PPTX 파일만 지원합니다.");
      return;
    }
    selectedFile = file;
    els.uploadContent.hidden = true;
    els.fileInfo.hidden = false;
    els.fileName.textContent = file.name;
    updateButtons();
  }

  function clearFile(e) {
    e.stopPropagation();
    selectedFile = null;
    els.fileInput.value = "";
    els.uploadContent.hidden = false;
    els.fileInfo.hidden = true;
    updateButtons();
  }

  function updateButtons() {
    var enabled = selectedFile !== null;
    els.btnAnalyze.disabled = !enabled;
    els.btnProcess.disabled = !enabled;
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

  // --- 파이프라인 실행 ---
  function runPipeline(mode) {
    if (!selectedFile) return;

    // UI 상태 변경
    els.btnAnalyze.disabled = true;
    els.btnProcess.disabled = true;
    els.progressSection.hidden = false;
    els.resultSection.hidden = true;

    var stageText = mode === "analyze" ? "분석" : "분석 + 변환";
    setProgress(0, stageText + " 시작...");

    var formData = new FormData();
    formData.append("file", selectedFile);

    var slides = els.slidesInput.value.trim();
    if (slides) {
      formData.append("slides", slides);
    }

    var endpoint = mode === "analyze" ? "/api/analyze" : "/api/process";
    setProgress(-1, stageText + " 진행 중...");

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
        els.progressSection.hidden = true;

        if (!result.ok) {
          showError(result.data.error || "알 수 없는 오류가 발생했습니다.");
          return;
        }

        if (mode === "analyze") {
          showAnalyzeResult(result.data);
        } else {
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

  // --- 결과 표시 ---
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

  function showError(message) {
    els.resultSection.hidden = false;
    els.resultContent.innerHTML =
      '<div class="result-summary error">' + message + "</div>";
  }
})();
