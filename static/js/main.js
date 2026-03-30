// Page unload cleanup - references variables from all tab modules
window.addEventListener('beforeunload', function() {
    if (typeof currentTaskId !== 'undefined' && currentTaskId && typeof isConverting !== 'undefined' && isConverting) {
        navigator.sendBeacon(`/api/cancel/${currentTaskId}`, new FormData());
    }
    if (typeof csvTaskId !== 'undefined' && csvTaskId && typeof isExtracting !== 'undefined' && isExtracting) {
        navigator.sendBeacon(`/api/csv/cleanup/${csvTaskId}`, new FormData());
    }
    if (typeof docxCsvTaskId !== 'undefined' && docxCsvTaskId && typeof isDocxCsvExtracting !== 'undefined' && isDocxCsvExtracting) {
        navigator.sendBeacon(`/api/docx-csv/cleanup/${docxCsvTaskId}`, new FormData());
    }
});
