window.editorInterop = {
    instance: null,
    dirty: false,

    _beforeUnloadHandler: function (e) {
        if (window.editorInterop.dirty) {
            e.preventDefault();
            e.returnValue = '';
        }
    },

    create: function (elementId, initialHtml) {
        const el = document.getElementById(elementId);
        if (!el) return;

        if (this.instance) {
            this.instance.destruct();
            this.instance = null;
        }

        this.dirty = false;

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

        // Track changes after initial content is set
        this.instance.events.on('change', () => {
            this.dirty = true;
        });

        window.addEventListener('beforeunload', this._beforeUnloadHandler);
    },

    getHtml: function () {
        if (!this.instance) return '';
        return this.instance.value;
    },

    clearDirty: function () {
        this.dirty = false;
    },

    dispose: function () {
        this.dirty = false;
        window.removeEventListener('beforeunload', this._beforeUnloadHandler);
        if (this.instance) {
            this.instance.destruct();
            this.instance = null;
        }
    }
};
