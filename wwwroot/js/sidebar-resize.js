(function () {
    function init() {
        const resizer = document.getElementById('sidebar-resizer');
        const sidebar = document.getElementById('app-sidebar');
        if (!resizer || !sidebar) return;
        if (resizer._initialized) return;
        resizer._initialized = true;

        let startX, startWidth;

        resizer.addEventListener('mousedown', function (e) {
            startX = e.clientX;
            startWidth = sidebar.offsetWidth;
            resizer.classList.add('dragging');
            document.body.style.cursor = 'col-resize';
            document.body.style.userSelect = 'none';

            function onMouseMove(e) {
                const newWidth = startWidth + (e.clientX - startX);
                const clamped = Math.max(180, Math.min(newWidth, window.innerWidth * 0.5));
                sidebar.style.width = clamped + 'px';
            }

            function onMouseUp() {
                resizer.classList.remove('dragging');
                document.body.style.cursor = '';
                document.body.style.userSelect = '';
                document.removeEventListener('mousemove', onMouseMove);
                document.removeEventListener('mouseup', onMouseUp);
            }

            document.addEventListener('mousemove', onMouseMove);
            document.addEventListener('mouseup', onMouseUp);
            e.preventDefault();
        });
    }

    // Initialize immediately (script is at bottom of body, DOM is ready)
    init();

    // Also retry after Blazor enhances the page
    document.addEventListener('DOMContentLoaded', init);
    new MutationObserver(init).observe(document.body, { childList: true, subtree: true });
})();
