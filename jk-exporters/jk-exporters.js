function Exporter() {
	// initialization
    includeSupport();
    includeExporters();

    function includeExporters() {
        var exporterRefs = [
            "./jk-exporters/excel-exporter/excel-exporter.js"
        ];

        includeSupportRef(exporterRefs);
    }

    function includeSupport() {

        // Include support files if ActiveX isn't present
        if (!(jk.hasActiveX)) {
            var fileRefs = [
                "./jk-exporters/support/Blob.js",
                "./jk-exporters/support/FileSaver.js"
            ];

            includeSupportRef(fileRefs);
        }
    }
}



function includeSupportRef(fileRefs) {
    $(fileRefs).each(function (i, fileRef) {
        var ref = document.createElement('script');
        ref.setAttribute("type", "text/javascript");
        ref.setAttribute("src", fileRef);

        document.getElementsByTagName("head")[0].appendChild(ref);
    });
}

function JK(){
    if (!(this instanceof JK)) return new JK();
}
jk = new JK();

JK.prototype.hasActiveX = !!window.ActiveXObject;

JK.prototype.exporters = new Exporter();
