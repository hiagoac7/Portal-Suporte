(function pauseAnimsWhenHidden(){
    if (window.__sdVisReady) return;
    window.__sdVisReady = true;
    function sync(){
        var h = document.documentElement;
        if (document.hidden) h.classList.add('sd-anim-paused');
        else h.classList.remove('sd-anim-paused');
    }
    document.addEventListener('visibilitychange', sync);
    sync();
})();