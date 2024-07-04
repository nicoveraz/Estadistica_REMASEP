window.onload = function() {
    document.getElementById("analyze").disabled = true;
};

function fileSelected() {
  document.getElementById("analyze").disabled = false;
}

function analyze(data) {
  // Your analysis code here
  alert("Analysis complete!");
}
