(function injectScrollbarCSS() {
    // Guard de instância única: o Streamlit re-renderiza este bloco a cada
    // rerun. Sem o guard, cada clique criava um novo MutationObserver +
    // setInterval que se acumulavam pela sessão, deixando o app cada vez
    // mais lento. Agora instala UMA vez só.
    if (window.__sdScrollbarReady) return;
    window.__sdScrollbarReady = true;

    const css = `
        * { scrollbar-width: thin; scrollbar-color: var(--ta550) transparent; }
        ::-webkit-scrollbar { width: 5px !important; height: 5px !important; }
        ::-webkit-scrollbar-track { background: rgba(var(--rgb-n1),.40) !important; border-radius: 99px !important; }
        ::-webkit-scrollbar-thumb {
            background: linear-gradient(180deg, var(--f-acc) 0%, var(--ta600) 100%) !important;
            border-radius: 99px !important;
            box-shadow: none;
        }
        ::-webkit-scrollbar-thumb:hover {
            background: var(--f-acc-l) !important;
            box-shadow: none;
        }
        ::-webkit-scrollbar-corner { background: transparent !important; }
    `;
    function applyToDoc(doc) {
        try {
            if (!doc.getElementById('_sd_scrollbar_fix')) {
                const s = doc.createElement('style');
                s.id = '_sd_scrollbar_fix';
                s.textContent = css;
                (doc.head || doc.documentElement).appendChild(s);
            }
        } catch(e) {}
    }
    function scanIframes() {
        document.querySelectorAll('iframe').forEach(function(f) {
            try { applyToDoc(f.contentDocument || f.contentWindow.document); } catch(e) {}
        });
    }
    // Debounce: o observer dispara a cada mutação do DOM (constantes no
    // Streamlit). Sem debounce, varria todos os iframes centenas de vezes
    // por segundo. Agora agrupa via requestAnimationFrame.
    let _scheduled = false;
    function scheduleScan() {
        if (_scheduled) return;
        _scheduled = true;
        requestAnimationFrame(function() { _scheduled = false; scanIframes(); });
    }
    applyToDoc(document);
    scanIframes();
    try {
        new MutationObserver(scheduleScan).observe(document.body, {childList:true, subtree:true});
    } catch(e) {}
})();