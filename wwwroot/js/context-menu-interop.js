window.contextMenuInterop = {
    dotNetRef: null,

    init: function (dotNetRef) {
        this.dotNetRef = dotNetRef;

        // Close menu on any click outside
        document.addEventListener('mousedown', function (e) {
            if (!e.target.closest('.context-menu')) {
                if (window.contextMenuInterop.dotNetRef) {
                    window.contextMenuInterop.dotNetRef.invokeMethodAsync('CloseMenu');
                }
            }
        });

        // Close menu on scroll (use capture to catch scrolling in any container)
        document.addEventListener('scroll', function () {
            if (window.contextMenuInterop.dotNetRef) {
                window.contextMenuInterop.dotNetRef.invokeMethodAsync('CloseMenu');
            }
        }, true);

        // Close on Escape key
        document.addEventListener('keydown', function (e) {
            if (e.key === 'Escape' && window.contextMenuInterop.dotNetRef) {
                window.contextMenuInterop.dotNetRef.invokeMethodAsync('CloseMenu');
            }
        });
    },

    // Nudge menu into viewport if it overflows edges
    adjustPosition: function () {
        var menu = document.querySelector('.context-menu');
        if (!menu) return;
        var rect = menu.getBoundingClientRect();
        if (rect.right > window.innerWidth) {
            menu.style.left = (window.innerWidth - rect.width - 8) + 'px';
        }
        if (rect.bottom > window.innerHeight) {
            menu.style.top = (window.innerHeight - rect.height - 8) + 'px';
        }
    }
};
