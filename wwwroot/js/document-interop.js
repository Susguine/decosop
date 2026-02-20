window.documentInterop = {
    // Load a PDF preview into an iframe element
    loadPdfPreview: function (frameId, url) {
        return new Promise(function (resolve) {
            var frame = document.getElementById(frameId);
            if (!frame) { resolve(); return; }

            frame.onload = function () { resolve(); };
            frame.onerror = function () { resolve(); };
            frame.src = url;

            // Fallback resolve after 30s in case onload doesn't fire
            setTimeout(resolve, 30000);
        });
    },

    // Print a PDF by opening it in a new tab and triggering the browser's print dialog
    printPdf: function (url) {
        var printWindow = window.open(url, '_blank');
        if (!printWindow) return;
        // Once the PDF loads in the new tab, trigger print
        printWindow.addEventListener('load', function () {
            setTimeout(function () { printWindow.print(); }, 500);
        });
    },

    // Print a file by opening a new window
    printFile: function (url, contentType, title) {
        var printWindow = window.open('', '_blank');
        if (!printWindow) return;

        if (contentType.startsWith('image/')) {
            printWindow.document.write(
                '<html><head><title>' + title + '</title>' +
                '<style>body{margin:0;display:flex;justify-content:center;align-items:center;min-height:100vh}' +
                'img{max-width:100%;max-height:100vh;object-fit:contain}</style></head>' +
                '<body><img src="' + url + '" onload="window.print()" /></body></html>'
            );
        } else if (contentType === 'application/pdf') {
            printWindow.document.write(
                '<html><head><title>' + title + '</title></head>' +
                '<body style="margin:0"><iframe src="' + url + '" ' +
                'style="width:100%;height:100vh;border:none" ' +
                'onload="setTimeout(function(){window.frames[0].focus();window.frames[0].print()},500)"></iframe></body></html>'
            );
        } else {
            printWindow.location.href = url;
        }
        printWindow.document.close();
    }
};
