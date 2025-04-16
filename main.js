let pyodide = null;
let inputFile = null;
let targetFile = null;
let downloadBlobUrl = null;
let cleanupTimeout = null;

async function loadPyodideAndPackages() {
  pyodide = await loadPyodide();
  await pyodide.loadPackage(["pandas", "openpyxl"]);
  document.getElementById("status").innerText = "Pyodide loaded!";
}

document.getElementById("inputFile").addEventListener("change", (e) => {
  inputFile = e.target.files[0];
});
document.getElementById("targetFile").addEventListener("change", (e) => {
  targetFile = e.target.files[0];
});

document.getElementById("runScript").addEventListener("click", async () => {
  if (!inputFile || !targetFile) {
    alert("Please upload both input and target Excel files.");
    return;
  }

  const selectedScript = document.getElementById("scriptSelect").value;

  try {
    document.getElementById("status").innerText = "Running script...";
    const scriptResponse = await fetch(selectedScript);
    const scriptText = await scriptResponse.text();

    const inputArrayBuffer = await inputFile.arrayBuffer();
    const targetArrayBuffer = await targetFile.arrayBuffer();

    pyodide.FS.writeFile("input.xlsx", new Uint8Array(inputArrayBuffer));
    pyodide.FS.writeFile("target.xlsx", new Uint8Array(targetArrayBuffer));

    let outputLog = await pyodide.runPythonAsync(`
import sys, io

sys.stdout = io.StringIO()
sys.stderr = io.StringIO()

${scriptText}

stdout = sys.stdout.getvalue()
stderr = sys.stderr.getvalue()
stdout + "\\n" + stderr
`);
    document.getElementById("pyOutput").innerText = outputLog;

    const updatedFile = pyodide.globals.get("output").toJs();
    const blob = new Blob([updatedFile], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    if (downloadBlobUrl) URL.revokeObjectURL(downloadBlobUrl);
    downloadBlobUrl = URL.createObjectURL(blob);

    const downloadBtn = document.getElementById("downloadFile");
    downloadBtn.href = downloadBlobUrl;
    downloadBtn.download = "updated_target.xlsx";
    downloadBtn.querySelector("button").disabled = false;
    document.getElementById("status").innerText =
      "Download ready. File will auto-delete in 5 minutes.";

    if (cleanupTimeout) clearTimeout(cleanupTimeout);
    cleanupTimeout = setTimeout(() => {
      try {
        pyodide.FS.unlink("input.xlsx");
        pyodide.FS.unlink("target.xlsx");
        URL.revokeObjectURL(downloadBlobUrl);
        document.getElementById("downloadFile").querySelector("button").disabled = true;
        document.getElementById("status").innerText = "Temporary files deleted after 5 minutes.";
        inputFile = null;
        targetFile = null;
        downloadBlobUrl = null;
      } catch (err) {
        console.warn("Cleanup error:", err);
      }
    }, 5 * 60 * 1000);

  } catch (err) {
    console.error("Error running script:", err);
    document.getElementById("pyOutput").innerText = err.toString();
    document.getElementById("status").innerText = "Python script failed.";
  }
});

loadPyodideAndPackages();

