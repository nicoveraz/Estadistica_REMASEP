document.getElementById("file-input").addEventListener("change", function() {
    document.getElementById("analyze").disabled = !this.files.length;
});

function analyze(data) {
  // Your analysis code here
  alert("Analysis complete!");
}
