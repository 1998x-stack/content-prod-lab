/* Content Production Lab · archive renderer */
(function(){
  "use strict";
  const $ = s => document.querySelector(s);
  const esc = s => String(s).replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
  const fmtSize = n => {
    if (n == null) return "";
    if (n < 1024) return n + " B";
    if (n < 1048576) return (n/1024).toFixed(0) + " KB";
    return (n/1048576).toFixed(1) + " MB";
  };
  const sealChar = name => {
    const m = String(name).match(/[\u4e00-\u9fffA-Za-z0-9]/);
    return m ? m[0] : "文";
  };
  const EXT_COLOR = { pdf:"#A83A2A", docx:"#2F5D8A", pptx:"#B0752B", xlsx:"#3C7A4E", html:"#5E6B62" };
  const extOf = it => it.ext || "pdf";
  const href = rel => "../" + rel.split("/").map(encodeURIComponent).join("/");

  const seal = (ch, cls) =>
    '<span class="seal' + (cls ? " " + cls : "") + '" aria-hidden="true">' + esc(ch) + "</span>";

  const fileRow = it => {
    const col = EXT_COLOR[extOf(it)] || "#5E6B62";
    return '<div class="file rv">' +
      seal(sealChar(it.name), "reg") +
      '<div class="info">' +
        '<div class="ttl">' + esc(it.name) + "</div>" +
        '<div class="sub">' +
          '<span class="chip' + (extOf(it)==="pdf" ? " on" : "") + '" style="color:' + col + ";border-color:" + col + '">'
            + "<small>" + esc(extOf(it).toUpperCase()) + "</small>" + (it.date ? " &middot; " + esc(it.date) : "") + "</span>" +
          '<span class="pmeta">' + fmtSize(it.size) + "</span>" +
        "</div>" +
        '<div class="ppath">' + esc(it.rel) + "</div>" +
      "</div>" +
      '<a class="open" href="' + href(it.rel) + '" target="_blank" rel="noopener" title="打开文件">' +
        seal("开") + "打开" + "</a>" +
    "</div>";
  };

  async function main(){
    const page = document.body.dataset.page;
    let data;
    try {
      data = await (await fetch("assets/data.json")).json();
    } catch(e){
      document.title = "载入失败";
      if ($("main")) $("main").innerHTML = '<div class="empty">无法载入数据清单，请确认部署于 /site 目录。</div>';
      return;
    }
    if ($("#gen")) $("#gen").textContent = data.generated;
    if ($("#gen2")) $("#gen2").textContent = data.generated;
    document.querySelectorAll("[data-stat='divs']").forEach(el=>el.textContent = data.divisions.length);
    document.querySelectorAll("[data-stat='files']").forEach(el=>el.textContent = data.total);

    if (page === "index") renderIndex(data);
    else if (page === "detail") renderDetail(data);
  }

  function renderIndex(data){
    const box = $("#index");
    if (!box) return;
    box.innerHTML = data.divisions.map((d,i) =>
      '<section class="div rv" style="animation-delay:' + (i*0.05) + 's">' +
        '<div class="col-seal">' + seal(sealChar(d.name), "reg") +
          '<span class="cnt">' + d.total + " 件</span></div>" +
        '<div class="body">' +
          '<div class="name"><a href="detail.html" title="' + esc(d.name) + '">' + esc(d.name) + "</a></div>" +
          '<div class="blurb">' + esc(d.blurb) + "</div>" +
          '<div class="folds">' +
            d.folders.map(f =>
              '<a class="fold" href="detail.html?p=' + encodeURIComponent(f.path) + '">' + esc(f.name) +
              '<span class="cnt">' + f.count + "</span></a>").join("") +
          "</div>" +
        "</div>" +
      "</section>").join("");
  }

  function renderDetail(data){
    const p = new URLSearchParams(location.search).get("p");
    let folder = null, div = null;
    for (const d of data.divisions){
      for (const f of d.folders){ if (f.path === p){ folder = f; div = d; break; } }
      if (folder) break;
    }
    const head = $("#head"), host = $("#host");
    if (!folder){
      head.innerHTML = "<h1>专题不存在</h1>";
      host.innerHTML = '<div class="empty">未找到对应专题，请从<a href="index.html">总览</a>进入。</div>';
      return;
    }
    document.title = folder.name + " · Content Production Lab";
    $("#fname").textContent = folder.name;
    $("#divname").textContent = div.name;
    if (folder.count) $("#fcount").textContent = folder.count + " 件";
    head.style.setProperty("--sealchar", sealChar(folder.name));

    const groups = {};
    folder.files.forEach(it => {
      let relp = it.rel.startsWith(folder.path + "/") ? it.rel.slice(folder.path.length + 1) : it.rel;
      const seg = relp.split("/");
      const g = seg.length > 1 ? seg[0] : folder.name;
      (groups[g] = groups[g] || []).push(it);
    });
    const keys = Object.keys(groups).sort(
      (a,b) => ((a===folder.name?0:1)-(b===folder.name?0:1)) || a.localeCompare(b, "zh"));

    host.innerHTML = keys.map(g =>
      '<div class="group"><div class="ghead">' + esc(g) + "</div>" +
      groups[g].map(fileRow).join("") + "</div>").join("");
    // stamp the seal char onto the detail masthead
    const sh = head.querySelector(".seal");
    if (sh) sh.textContent = sealChar(folder.name);
  }

  document.addEventListener("DOMContentLoaded", main);
})();