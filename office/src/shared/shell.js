// Application frame shared by Docs, Sheets and Slides: title bar with quick
// access buttons, ribbon, File backstage, status bar, Claude pane, dialogs and
// the open/save lifecycle.

import { icon } from './icons.js';
import { Ribbon } from './ribbon.js';
import { ClaudePane } from './claude-pane.js';

export const APPS = {
  docs: { title: 'Rea Docs', accent: '#2f5aa8', kind: 'Document', mark: 'appDocs', newLabel: 'Blank document' },
  sheets: { title: 'Rea Sheets', accent: '#1b7a4b', kind: 'Workbook', mark: 'appSheets', newLabel: 'Blank workbook' },
  slides: { title: 'Rea Slides', accent: '#c2532d', kind: 'Presentation', mark: 'appSlides', newLabel: 'Blank presentation' },
};

export const h = (tag, cls, html) => {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (html != null) e.innerHTML = html;
  return e;
};

export function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

export function toast(message, kind = 'info') {
  const t = h('div', `toast toast-${kind}`, escapeHtml(message));
  document.body.append(t);
  requestAnimationFrame(() => t.classList.add('show'));
  setTimeout(() => {
    t.classList.remove('show');
    setTimeout(() => t.remove(), 300);
  }, 2600);
}

// Small modal dialog. fields: [{name, label, type, value, options}]
export function dialog({ title, body, fields = [], ok = 'OK', cancel = 'Cancel', wide = false }) {
  return new Promise((resolve) => {
    const back = h('div', 'modal-back');
    const box = h('form', `modal${wide ? ' wide' : ''}`);
    box.append(h('div', 'modal-title', escapeHtml(title)));
    if (body) box.append(typeof body === 'string' ? h('div', 'modal-body', body) : body);
    const inputs = {};
    for (const f of fields) {
      const row = h('label', 'modal-field');
      row.append(h('span', '', escapeHtml(f.label)));
      let input;
      if (f.type === 'select') {
        input = h('select');
        for (const o of f.options) {
          const opt = h('option', '', escapeHtml(o.label));
          opt.value = o.value;
          input.append(opt);
        }
      } else if (f.type === 'textarea') {
        input = h('textarea');
        input.rows = f.rows || 4;
      } else {
        input = h('input');
        input.type = f.type || 'text';
        if (f.min != null) input.min = f.min;
        if (f.max != null) input.max = f.max;
        if (f.placeholder) input.placeholder = f.placeholder;
      }
      if (f.value != null) input.value = f.value;
      inputs[f.name] = input;
      row.append(input);
      box.append(row);
    }
    const actions = h('div', 'modal-actions');
    const okBtn = h('button', 'btn primary', escapeHtml(ok));
    okBtn.type = 'submit';
    actions.append(okBtn);
    if (cancel) {
      const c = h('button', 'btn', escapeHtml(cancel));
      c.type = 'button';
      c.onclick = () => done(null);
      actions.append(c);
    }
    box.append(actions);
    back.append(box);
    document.body.append(back);
    const first = Object.values(inputs)[0];
    (first || okBtn).focus();
    first?.select?.();
    function done(v) {
      back.remove();
      resolve(v);
    }
    box.onsubmit = (e) => {
      e.preventDefault();
      const out = {};
      for (const [k, i] of Object.entries(inputs)) out[k] = i.value;
      done(out);
    };
    back.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') done(null);
    });
  });
}

export async function createShell(opts) {
  const meta = APPS[opts.app];
  const info = await window.office.launchInfo();
  document.documentElement.dataset.platform = info.platform;
  document.documentElement.style.setProperty('--accent', meta.accent);
  document.title = meta.title;

  const state = {
    path: info.file?.path || null,
    name: info.file?.name || `${meta.kind}1`,
    dirty: false,
  };

  const root = h('div', `app app-${opts.app}`);
  const titlebar = h('header', 'titlebar');
  const qa = h('div', 'qa');
  const mark = h('div', 'titlebar-mark', icon(meta.mark, 20));
  const qaButtons = [
    { icon: 'save', title: 'Save (Ctrl+S)', run: () => shell.save() },
    { icon: 'undo', title: 'Undo (Ctrl+Z)', run: () => opts.undo?.() },
    { icon: 'redo', title: 'Redo (Ctrl+Y)', run: () => opts.redo?.() },
  ];
  for (const b of qaButtons) {
    const btn = h('button', 'qa-btn', icon(b.icon, 16));
    btn.title = b.title;
    btn.onmousedown = (e) => e.preventDefault();
    btn.onclick = b.run;
    qa.append(btn);
  }
  const titleText = h('div', 'titlebar-title');
  const search = h('div', 'titlebar-search', `${icon('sparkle', 14)}<input placeholder="Ask Claude or search commands" spellcheck="false">`);
  titlebar.append(mark, qa, titleText, search, h('div', 'titlebar-spacer'));

  const ribbonHost = h('div');
  const tabs = [...opts.tabs, ...(opts.claudeTab ? [opts.claudeTab] : [])];
  const workspaceRow = h('div', 'workspace-row');
  const workspace = h('main', 'workspace');
  const paneHost = h('aside', 'claude-host');
  workspaceRow.append(workspace, paneHost);
  const status = h('footer', 'statusbar');
  const statusLeft = h('div', 'status-left');
  const statusRight = h('div', 'status-right');
  status.append(statusLeft, statusRight);
  root.append(titlebar, ribbonHost, workspaceRow, status);
  document.body.append(root);

  const shell = {
    info,
    state,
    root,
    workspace,
    statusLeft,
    statusRight,
    toast,
    dialog,
    get ribbon() {
      return ribbon;
    },
    setTitle() {
      titleText.textContent = `${state.name}${state.dirty ? ' •' : ''} — ${meta.title}`;
      document.title = `${state.name} - ${meta.title}`;
      window.office.setTitle(document.title);
    },
    markDirty() {
      if (!state.dirty) {
        state.dirty = true;
        window.office.setDirty(true);
        shell.setTitle();
      }
    },
    markClean() {
      state.dirty = false;
      window.office.setDirty(false);
      shell.setTitle();
    },
    async save({ saveAs = false, ext } = {}) {
      const payload = await opts.getSaveData();
      const baseName = state.name.replace(/\.[^.]+$/, '');
      const res = await window.office.save({
        app: opts.app,
        payload,
        path: ext ? null : state.path,
        saveAs: saveAs || !!ext,
        suggestedName: ext ? `${baseName}.${ext}` : state.path || `${baseName}.${opts.defaultExt}`,
      });
      if (!res) return false;
      if (res.exported) {
        toast(`Exported ${res.name}`);
        return true;
      }
      state.path = res.path;
      state.name = res.name;
      shell.markClean();
      toast(`Saved ${res.name}`);
      return true;
    },
    async exportPdf() {
      document.body.classList.add('printing');
      opts.beforePrint?.();
      try {
        await window.office.exportPdf({ suggestedName: state.name.replace(/\.[^.]+$/, ''), ...(opts.pdfOptions?.() || {}) });
      } finally {
        document.body.classList.remove('printing');
        opts.afterPrint?.();
      }
    },
    async print() {
      document.body.classList.add('printing');
      opts.beforePrint?.();
      await window.office.print();
      setTimeout(() => {
        document.body.classList.remove('printing');
        opts.afterPrint?.();
      }, 500);
    },
    openBackstage(section = 'home') {
      backstage.open(section);
    },
    refresh() {
      ribbon.refresh();
    },
  };

  const ribbon = new Ribbon(ribbonHost, { tabs, accent: meta.accent, onFile: () => shell.openBackstage() });
  const backstage = new Backstage(shell, opts, meta);
  root.append(backstage.node);

  if (opts.claude) {
    shell.claude = new ClaudePane(paneHost, { ...opts.claude, shell });
    const input = search.querySelector('input');
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && input.value.trim()) {
        shell.claude.show();
        shell.claude.ask(input.value.trim());
        input.value = '';
      }
      if (e.key === 'Escape') input.blur();
    });
  } else {
    search.remove();
  }

  window.office.onSaveAndClose(async () => {
    if (await shell.save()) window.office.close();
  });

  document.addEventListener('keydown', (e) => {
    const mod = e.ctrlKey || e.metaKey;
    if (e.key === 'F1' && mod) {
      e.preventDefault();
      ribbon.toggleCollapsed();
      return;
    }
    if (!mod) return;
    const k = e.key.toLowerCase();
    if (k === 's') {
      e.preventDefault();
      shell.save({ saveAs: e.shiftKey });
    } else if (k === 'o') {
      e.preventDefault();
      window.office.openDialog(opts.app);
    } else if (k === 'n') {
      e.preventDefault();
      window.office.newWindow(opts.app);
    } else if (k === 'p') {
      e.preventDefault();
      shell.print();
    } else if (k === 'j' || (k === 'k' && e.shiftKey)) {
      e.preventDefault();
      shell.claude?.toggle();
    }
  });

  shell.setTitle();
  return shell;
}

class Backstage {
  constructor(shell, opts, meta) {
    this.shell = shell;
    this.opts = opts;
    this.meta = meta;
    this.node = h('div', 'backstage');
    this.nav = h('nav', 'bs-nav');
    this.content = h('section', 'bs-content');
    this.node.append(this.nav, this.content);
    const items = [
      { id: 'back', label: '', icon: 'back', run: () => this.close() },
      { id: 'home', label: 'Home', icon: 'home' },
      { id: 'new', label: 'New', icon: 'newFile' },
      { id: 'open', label: 'Open', icon: 'open' },
      '-',
      { id: 'save', label: 'Save', icon: 'save', run: () => this.close() || shell.save() },
      { id: 'saveas', label: 'Save As', icon: 'saveAs', run: () => this.close() || shell.save({ saveAs: true }) },
      { id: 'export', label: 'Export', icon: 'export' },
      { id: 'print', label: 'Print', icon: 'print', run: () => this.close() || shell.print() },
      '-',
      { id: 'options', label: 'Options', icon: 'settings' },
      { id: 'close', label: 'Close', icon: 'close', run: () => window.office.close() },
    ];
    this.buttons = {};
    for (const it of items) {
      if (it === '-') {
        this.nav.append(h('div', 'bs-sep'));
        continue;
      }
      const b = h('button', `bs-item${it.id === 'back' ? ' bs-back' : ''}`, `${icon(it.icon, 18)}<span>${it.label}</span>`);
      b.onclick = () => (it.run ? it.run() : this.open(it.id));
      this.buttons[it.id] = b;
      this.nav.append(b);
    }
    this.node.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') this.close();
    });
  }

  close() {
    this.node.classList.remove('open');
  }

  async open(section) {
    this.node.classList.add('open');
    for (const [id, b] of Object.entries(this.buttons)) b.classList.toggle('active', id === section);
    this.content.innerHTML = '';
    const c = this.content;
    const appTiles = () => {
      const row = h('div', 'bs-tiles');
      for (const [id, m] of Object.entries(APPS)) {
        const t = h('button', `bs-tile bs-tile-${id}`, `<div class="bs-tile-preview">${icon(m.mark, 40)}</div><span>${m.newLabel}</span>`);
        t.onclick = () => {
          this.close();
          window.office.newWindow(id);
        };
        row.append(t);
      }
      return row;
    };
    if (section === 'home' || section === 'new') {
      c.append(h('h1', '', section === 'home' ? 'Good to see you' : 'New'));
      c.append(h('h2', '', 'Start something new'));
      c.append(appTiles());
      c.append(h('h2', '', 'Recent'));
      c.append(await recentList(() => this.close()));
    } else if (section === 'open') {
      c.append(h('h1', '', 'Open'));
      const browse = h('button', 'bs-browse', `${icon('open', 28)}<span>Browse…</span>`);
      browse.onclick = () => {
        this.close();
        window.office.openDialog(null);
      };
      c.append(browse, h('h2', '', 'Recent'), await recentList(() => this.close()));
    } else if (section === 'export') {
      c.append(h('h1', '', 'Export'));
      const list = h('div', 'bs-exports');
      const exports = [{ label: 'PDF document', detail: 'Fixed layout that looks the same everywhere.', icon: 'pdf', run: () => this.shell.exportPdf() }, ...(this.opts.exports || [])];
      for (const ex of exports) {
        const b = h('button', 'bs-export', `${icon(ex.icon || 'export', 32)}<div><b>${escapeHtml(ex.label)}</b><span>${escapeHtml(ex.detail || '')}</span></div>`);
        b.onclick = () => {
          this.close();
          if (ex.run) ex.run();
          else this.shell.save({ ext: ex.ext });
        };
        list.append(b);
      }
      c.append(list);
    } else if (section === 'options') {
      c.append(h('h1', '', 'Options'));
      c.append(await settingsForm());
    }
  }
}

async function recentList(onOpen) {
  const list = h('div', 'bs-recent');
  const items = await window.office.recent();
  if (!items.length) list.append(h('div', 'bs-empty', 'Files you open or save show up here.'));
  for (const r of items) {
    const m = APPS[r.app] || APPS.docs;
    const name = r.path.split(/[\\/]/).pop();
    const row = h('button', 'bs-recent-row', `${icon(m.mark, 22)}<div><b>${escapeHtml(name)}</b><span>${escapeHtml(r.path)}</span></div><time>${new Date(r.at).toLocaleDateString()}</time>`);
    row.style.setProperty('--accent', m.accent);
    row.onclick = () => {
      onOpen?.();
      window.office.openPath(r.path);
    };
    list.append(row);
  }
  return list;
}

export async function settingsForm(onSaved) {
  const s = await window.office.settings.get();
  const form = h('form', 'settings-form');
  form.innerHTML = `
    <h2>${icon('sparkle', 18)} Claude</h2>
    <p class="muted">Claude runs through your own Anthropic API key. The key is stored encrypted on this computer and is only sent to Anthropic.</p>
    <label class="modal-field"><span>Anthropic API key</span>
      <input name="apiKey" type="password" autocomplete="off" placeholder="${s.hasKey ? '•••••••• (saved)' : 'sk-ant-…'}" ${s.keyFromEnv ? 'disabled' : ''}>
    </label>
    ${s.keyFromEnv ? '<p class="muted">Using the ANTHROPIC_API_KEY environment variable.</p>' : ''}
    <p><a href="#" data-link="https://console.anthropic.com/settings/keys">Get an API key from the Anthropic Console</a></p>
    <label class="modal-field"><span>Model</span><select name="model">${s.models.map((m) => `<option value="${m.id}" ${m.id === s.model ? 'selected' : ''}>${escapeHtml(m.label)}</option>`).join('')}</select></label>
    <label class="modal-field"><span>Thinking effort</span><select name="effort">${['low', 'medium', 'high', 'xhigh', 'max'].map((e) => `<option ${e === s.effort ? 'selected' : ''}>${e}</option>`).join('')}</select></label>
    <div class="modal-actions"><button class="btn primary" type="submit">Save settings</button>${s.hasKey && !s.keyFromEnv ? '<button class="btn" type="button" data-remove>Remove key</button>' : ''}</div>`;
  form.querySelector('[data-link]').onclick = (e) => {
    e.preventDefault();
    window.office.openExternal(e.target.dataset.link);
  };
  form.onsubmit = async (e) => {
    e.preventDefault();
    const fd = new FormData(form);
    const patch = { model: fd.get('model'), effort: fd.get('effort') };
    if (fd.get('apiKey')) patch.apiKey = fd.get('apiKey');
    const next = await window.office.settings.set(patch);
    toast('Settings saved');
    onSaved?.(next);
  };
  form.querySelector('[data-remove]')?.addEventListener('click', async () => {
    const next = await window.office.settings.set({ apiKey: '' });
    toast('API key removed');
    onSaved?.(next);
  });
  return form;
}
