window.editorInterop = {
    instance: null,

    create: function (elementId, initialHtml) {
        const el = document.getElementById(elementId);
        if (!el) return;

        if (this.instance) {
            this.instance.destruct();
            this.instance = null;
        }

        this.instance = Jodit.make('#' + elementId, {
            height: 'calc(100vh - 14rem)',
            toolbarSticky: false,
            showCharsCounter: false,
            showWordsCounter: false,
            showXPathInStatusbar: false,
            buttons: [
                'bold', 'italic', 'underline', 'strikethrough', '|',
                'ul', 'ol', '|',
                'font', 'fontsize', 'brush', '|',
                'align', 'indent', 'outdent', '|',
                'table', '|',
                'link', 'hr', '|',
                'undo', 'redo', '|',
                'eraser', 'fullsize'
            ],
            placeholder: 'Start writing your SOP here...',
            askBeforePasteHTML: false,
            askBeforePasteFromWord: false,
            defaultActionOnPaste: 'insert_clear_html'
        });

        if (initialHtml) {
            this.instance.value = initialHtml;
        }
    },

    getHtml: function () {
        if (!this.instance) return '';
        return this.instance.value;
    },

    dispose: function () {
        if (this.instance) {
            this.instance.destruct();
            this.instance = null;
        }
    }
};
