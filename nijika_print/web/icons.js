/* Locally drawn vector symbols, with no remote fonts or icon services. */
const iconPaths={
 printer:'<path d="M7 8V3h10v5M7 17H4a1 1 0 0 1-1-1v-6a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2v6a1 1 0 0 1-1 1h-3"/><path d="M7 14h10v7H7zM8 11h.01M10 17h4"/>',
 folder:'<path d="M3 6a2 2 0 0 1 2-2h5l2 3h7a2 2 0 0 1 2 2v10a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1z"/>',
 'folder-plus':'<path d="M3 6a2 2 0 0 1 2-2h5l2 3h7a2 2 0 0 1 2 2v10a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1zM12 10v7m-3.5-3.5h7"/>',
 plus:'<path d="M12 5v14M5 12h14"/>',minus:'<path d="M5 12h14"/>',maximize:'<rect x="5" y="5" width="14" height="14" rx="1"/>',close:'<path d="m6 6 12 12M18 6 6 18"/>',
 search:'<circle cx="10.5" cy="10.5" r="6.5"/><path d="m16 16 4 4"/>',refresh:'<path d="M20 7v5h-5M4 17v-5h5"/><path d="M5.4 7.5A8 8 0 0 1 19.8 12M4.2 12a8 8 0 0 0 14.4 4.5"/>',
 'arrow-up':'<path d="M12 19V5m-5 5 5-5 5 5"/>','arrow-down':'<path d="M12 5v14m-5-5 5 5 5-5"/>','arrow-right':'<path d="M4 12h16m-6-6 6 6-6 6"/>',
 chevron:'<path d="m8 10 4 4 4-4"/>','chevron-right':'<path d="m10 7 5 5-5 5"/>',check:'<path d="m5 12 4.5 4.5L19 7"/>',
 file:'<path d="M14 3H6a1 1 0 0 0-1 1v16a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V8zM14 3v5h5M9 12h6m-6 4h6"/>',
 documents:'<rect x="4" y="7" width="12" height="15" rx="2"/><path d="M9 3h10a2 2 0 0 1 2 2v13M8 12h4m-4 4h4"/>',
 image:'<rect x="3" y="3" width="18" height="18" rx="3"/><circle cx="8" cy="8" r="1"/><path d="m3 17 5-5 4 4 4-6 5 7"/>',
 trash:'<path d="M4 6h16M9 6V3h6v3M6 6l1 15h10l1-15M10 10v7m4-7v7"/>',
 sliders:'<path d="M3 6h4m4 0h10M3 12h11m4 0h3M3 18h3m4 0h11"/><rect x="7" y="4" width="4" height="4" rx="1"/><rect x="14" y="10" width="4" height="4" rx="1"/><rect x="6" y="16" width="4" height="4" rx="1"/>',
 help:'<circle cx="12" cy="12" r="9"/><path d="M9.5 8.5a2.5 2.5 0 0 1 5 .5c0 2-2.5 2-2.5 4M12 17h.01"/>',info:'<circle cx="12" cy="12" r="9"/><path d="M12 10v6m0-9h.01"/>',
 spark:Array.from({length:12},(_,i)=>{const a=i*Math.PI/6;const r=i%2?8:10;return `<path d="M${12+Math.cos(a)*3} ${12+Math.sin(a)*3}L${12+Math.cos(a)*r} ${12+Math.sin(a)*r}"/>`}).join('')
};
function icon(name){return `<svg class="icon" viewBox="0 0 24 24" aria-hidden="true">${iconPaths[name]||iconPaths.file}</svg>`}
function hydrateIcons(root=document){root.querySelectorAll('[data-icon]').forEach(el=>el.innerHTML=icon(el.dataset.icon))}
hydrateIcons();
