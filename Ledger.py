import hashlib
import json
from io import BytesIO

import pandas as pd
import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode

ORDER = [
    'Sales - net', 'Cost of sales', 'Selling and distribution expenses',
    'Administrative expenses', 'Other operating incomes',
    'Other operating expenses', 'Finance Costs', 'Taxation',
]
MAIN, SUB, CODE = 'Current Main Class', 'Current Sub Class', 'Acct Code'
AMOUNT = 'Deb./Cred. (LC)'


def normalize_code(series):
    return series.astype('string').str.strip().str.replace(r'\.0$', '', regex=True).replace('', pd.NA)


@st.cache_data
def load_data(content):
    with pd.ExcelFile(BytesIO(content)) as workbook:
        missing = {'General Ledger', 'Acct class'} - set(workbook.sheet_names)
        if missing:
            raise ValueError('Missing worksheets: ' + ', '.join(sorted(missing)))
        gl = pd.read_excel(workbook, sheet_name='General Ledger')
        mapping = pd.read_excel(workbook, sheet_name='Acct class')
    for table, required, label in [
        (gl, [CODE, 'Acct Name', 'Posting Date', AMOUNT], 'General Ledger'),
        (mapping, ['Account code', MAIN, SUB, 'Title'], 'Acct class'),
    ]:
        table.columns = table.columns.astype(str).str.strip()
        missing = set(required) - set(table.columns)
        if missing:
            raise ValueError(f'{label}: missing columns: {", ".join(sorted(missing))}')
    gl[CODE] = normalize_code(gl[CODE])
    mapping['Account code'] = normalize_code(mapping['Account code'])
    if gl[CODE].isna().any():
        raise ValueError('General Ledger contains blank account codes. Correct them before uploading.')
    mapping = mapping[['Account code', MAIN, SUB, 'Title']].dropna(subset=['Account code'])
    for col in [MAIN, SUB, 'Title']:
        mapping[col] = mapping[col].astype('string').str.strip().replace('', pd.NA)
    mapping = mapping.drop_duplicates()
    duplicates = mapping.loc[mapping['Account code'].duplicated(False), 'Account code'].unique()
    if len(duplicates):
        raise ValueError('Conflicting account classifications for codes: ' + ', '.join(duplicates[:10]))
    dates = pd.to_datetime(gl['Posting Date'], errors='coerce')
    if dates.isna().any():
        raise ValueError('General Ledger contains missing or invalid posting dates. Correct them before uploading.')
    amounts = pd.to_numeric(gl[AMOUNT], errors='coerce')
    if amounts.isna().any():
        raise ValueError('General Ledger contains missing or nonnumeric ledger amounts.')
    gl['Posting Date'], gl[AMOUNT] = dates, amounts
    # Remove any old classifications so the mapping sheet remains authoritative.
    gl = gl.drop(columns=[MAIN, SUB, 'Title'], errors='ignore')
    df = gl.merge(mapping, left_on=CODE, right_on='Account code', how='left', validate='many_to_one')
    df[MAIN] = df[MAIN].fillna('Unclassified')
    df[SUB] = df[SUB].fillna('Unclassified')
    df['Title'] = df['Title'].fillna(df['Acct Name']).fillna(df[CODE])
    # One stable display title per code; drill-down always uses the account code.
    df['Title'] = df.groupby(CODE, sort=False)['Title'].transform('first')
    df['Year'] = df['Posting Date'].dt.year
    return df.loc[df['Posting Date'].dt.month.eq(9) & df['Year'].isin([2025, 2026])].copy()


def make_summary(df):
    pivot = df.pivot_table(index=[MAIN, SUB, CODE, 'Title'], columns='Year', values=AMOUNT,
                           aggfunc='sum', fill_value=0, observed=True)
    pivot = pivot.reindex(columns=[2026, 2025], fill_value=0).mul(-1)
    pivot.columns = ['2026', '2025']
    return pivot.reset_index()


def visible_rows(summary, expanded):
    rows, paths = [], []

    def append_row(part, path, label):
        old, new = float(part['2025'].sum()), float(part['2026'].sum())
        change = (new - old) / abs(old) if old != 0 else None
        marker = ('− ' if path in expanded else '+ ') if len(path) < 3 else ''
        if len(path) < 3:
            count = part[CODE].nunique()
            label = f'{label} ({count} {"account" if count == 1 else "accounts"})'
        rows.append({'Hierarchy': '    ' * (len(path) - 1) + marker + label,
                     '2026': new, '2025': old, '% Change': change,
                     '_level': len(path), '_path': json.dumps(path, ensure_ascii=False, separators=(',', ':')),
                     '_expanded': '', '_action': '', '_id': len(paths)})
        paths.append(path)

    main_values = sorted(summary[MAIN].unique(), key=lambda x: (ORDER.index(x) if x in ORDER else len(ORDER), x))
    for main in main_values:
        part = summary.loc[summary[MAIN].eq(main)]
        append_row(part, (main,), main)
        if (main,) not in expanded:
            continue
        for sub, subgroup in part.groupby(SUB, sort=True):
            append_row(subgroup, (main, sub), sub)
            if (main, sub) in expanded:
                for _, account in subgroup.sort_values(['Title', CODE]).iterrows():
                    append_row(account.to_frame().T, (main, sub, account[CODE]),
                               f"{account['Title']} [{account[CODE]}]")
    return pd.DataFrame(rows), paths


def run():
    st.set_page_config(layout='wide', page_title='September Financial Dashboard')
    st.title('September 2026 vs 2025 Ledger Dashboard')
    # Initialize new keys even when an existing session survives an app update.
    for key, default in [('expanded', set()), ('grid_revision', 0),
                         ('ledger_path', None), ('ledger_year', 2026)]:
        st.session_state.setdefault(key, default)
    upload = st.file_uploader('Upload an SAP Excel workbook containing the General Ledger and Acct class worksheets.', type=['xlsx'])
    if upload is None:
        st.info('Upload your Excel workbook to view the financial summary and ledger details.')
        return
    try:
        content = upload.getvalue()
        fingerprint = hashlib.sha256(content).hexdigest()
        if st.session_state.get('workbook_id') != fingerprint:
            st.session_state.update(workbook_id=fingerprint, expanded=set(), grid_revision=0, ledger_path=None)
        df = load_data(content)
        if df.empty:
            st.warning('No September transactions were found for 2025 or 2026.')
            return
        summary = make_summary(df)
        unclassified = df.loc[df[MAIN].eq('Unclassified'), CODE].nunique()
        if unclassified:
            st.warning(f'{unclassified} account(s) have no main classification and are shown under Unclassified.')
        st.subheader('1. September Profit and Loss Summary')
        st.caption('P&L values = ledger values × −1. Ledger details retain their original debit/credit signs. Percentage change uses the absolute prior-year balance; a zero prior-year balance is shown as N/A.')
        st.caption('Level 1: bold navy rows · Level 2: light blue rows · Level 3: indented white account rows')
        all_branches = {(m,) for m in summary[MAIN].unique()}
        all_branches.update(tuple(r) for r in summary[[MAIN, SUB]].drop_duplicates().itertuples(index=False, name=None))
        table, paths = visible_rows(summary, all_branches)
        gb = GridOptionsBuilder.from_dataframe(table)
        gb.configure_default_column(sortable=False, filter=False, resizable=True)
        gb.configure_column('_id', hide=True)
        gb.configure_column('_level', hide=True)
        gb.configure_column('_action', hide=True)
        gb.configure_column('_path', hide=True)
        gb.configure_column('_expanded', hide=True)
        gb.configure_column('Hierarchy', header_name='Main Class / Sub Class / Account', minWidth=450,
                            cellStyle={'whiteSpace': 'pre'},
                            headerComponent=JsCode("""class LevelHeader {
                init(p) {
                    this.el = document.createElement('div');
                    this.el.style.cssText = 'width:100%;display:flex;flex-direction:column;gap:8px;';
                    const title = document.createElement('span');
                    title.textContent = p.displayName;
                    this.el.appendChild(title);
                    const controls = document.createElement('div');
                    controls.style.cssText = 'display:flex;gap:6px;';
                    ['Level 1', 'Level 2', 'Level 3'].forEach((label, index) => {
                        const button = document.createElement('button');
                        button.type = 'button';
                        button.textContent = label;
                        button.title = ['Main classes / collapse all', 'Show sub classes', 'Show accounts / expand all'][index];
                        button.style.cssText = 'cursor:pointer;padding:4px 12px;border:1px solid #a7bbce;border-radius:5px;background:#ffffff;color:#17324d;font-size:13px;';
                        button.addEventListener('click', (event) => {
                            event.stopPropagation();
                            const opened = [];
                            p.api.forEachNode(node => {
                                if (node.data && node.data._level <= index) opened.push(node.data._path);
                            });
                            p.context.expanded = opened;
                            p.api.onFilterChanged();
                            p.api.refreshCells({columns: ['Hierarchy'], force: true});
                        });
                        controls.appendChild(button);
                    });
                    this.el.appendChild(controls);
                }
                getGui() { return this.el; }
            }"""),
                            cellRenderer=JsCode("""class HierarchyRenderer {
                init(p) {
                    this.el = document.createElement('div');
                    if (p.node.rowPinned) {
                        this.el.textContent = p.value;
                        return;
                    }
                    this.el.style.display = 'flex';
                    this.el.style.alignItems = 'center';
                    this.el.style.height = '100%';
                    this.el.style.paddingLeft = ((p.data._level - 1) * 22) + 'px';
                    const label = String(p.value || '').trim();
                    if (p.data._level < 3) {
                        const button = document.createElement('button');
                        button.type = 'button';
                        const isOpen = (p.context.expanded || []).includes(p.data._path);
                        button.textContent = isOpen ? '−' : '+';
                        button.title = isOpen ? 'Collapse category' : 'Expand category';
                        button.setAttribute('aria-label', button.title);
                        button.setAttribute('aria-expanded', String(isOpen));
                        button.style.cssText = 'cursor:pointer;width:26px;height:26px;margin-right:9px;border:1px solid currentColor;border-radius:5px;background:transparent;color:inherit;font-size:18px;line-height:20px;';
                        button.addEventListener('click', function(e) {
                            e.stopPropagation();
                            const key = p.data._path;
                            const opened = p.context.expanded || [];
                            if (opened.includes(key)) {
                                const prefix = JSON.parse(key);
                                p.context.expanded = opened.filter(function(item) {
                                    const path = JSON.parse(item);
                                    return !prefix.every((value, i) => path[i] === value);
                                });
                            } else {
                                p.context.expanded = opened.concat([key]);
                            }
                            p.api.onFilterChanged();
                            p.api.refreshCells({columns: ['Hierarchy'], force: true});
                        });
                        this.el.appendChild(button);
                    }
                    const text = document.createElement('span');
                    text.textContent = p.data._level < 3 ? label.slice(2) : label;
                    this.el.appendChild(text);
                }
                getGui() { return this.el; }
            }"""))
        money = JsCode("function(p) { return p.value == null ? '' : Number(p.value).toLocaleString('en-US', {minimumFractionDigits:2, maximumFractionDigits:2}); }")
        percent = JsCode("function(p) { return p.value == null ? 'N/A' : Number(p.value).toLocaleString('en-US', {style:'percent', minimumFractionDigits:2, maximumFractionDigits:2}); }")
        for col in ['2026', '2025']:
            gb.configure_column(col, type=['numericColumn'], valueFormatter=money)
        gb.configure_column('% Change', type=['numericColumn'], valueFormatter=percent)
        total_new, total_old = float(summary['2026'].sum()), float(summary['2025'].sum())
        gb.configure_grid_options(
            headerHeight=78,
            pinnedBottomRowData=[{'Hierarchy': 'Grand Total', '2026': total_new,
                                  '2025': total_old,
                                  '% Change': (total_new - total_old) / abs(total_old) if total_old else None}],
            context={'expanded': [json.dumps(path, ensure_ascii=False, separators=(',', ':'))
                                  for path in sorted(st.session_state.expanded)]},
            isExternalFilterPresent=JsCode('function() { return true; }'),
            doesExternalFilterPass=JsCode('''function(p) {
                const path = JSON.parse(p.data._path);
                for (let depth = 1; depth < path.length; depth++) {
                    if (!(p.context.expanded || []).includes(JSON.stringify(path.slice(0, depth)))) return false;
                }
                return true;
            }'''),
            onGridReady=JsCode('function(p) { p.api.onFilterChanged(); }'),
            onFirstDataRendered=JsCode('function(p) { p.api.onFilterChanged(); }'),
            onCellClicked=JsCode('''function(p) {
                if (!p.data || p.node.rowPinned || p.colDef.field === '_action') return;
                // Toggle buttons stop propagation; every other cell opens the ledger.
                // Keep one pending action and preserve the browser's expanded branches.
                p.api.forEachNode(function(node) { node.data._action = ''; });
                p.data._expanded = JSON.stringify(p.context.expanded || []);
                p.node.setDataValue('_action', 'ledger:' + p.colDef.field + ':' + Date.now());
            }'''),
            getRowStyle=JsCode("""function(p) {
                if (!p.data) return {};
                if (p.node.rowPinned) return {
                    backgroundColor: '#17324d', color: '#ffffff', fontWeight: '700',
                    fontSize: '14px', borderTop: '2px solid #a7bbce'
                };
                if (p.data._level === 1) return {
                    backgroundColor: '#17324d', color: '#ffffff',
                    fontWeight: '700', fontSize: '14px',
                    borderTop: '2px solid #ffffff'
                };
                if (p.data._level === 2) return {
                    backgroundColor: '#e7eff7', color: '#17324d',
                    fontWeight: '600', fontSize: '13px',
                    borderBottom: '1px solid #cedcea'
                };
                return {
                    backgroundColor: '#ffffff', color: '#374151',
                    fontWeight: '400', fontSize: '13px',
                    borderBottom: '1px solid #eef2f6'
                };
            }"""),
            getRowHeight=JsCode("function(p) { return p.data && p.data._level === 1 ? 44 : 36; }")
        )
        gb.configure_selection('single', use_checkbox=False)
        response = AgGrid(table, gridOptions=gb.build(), height=480, theme='streamlit',
                          allow_unsafe_jscode=True, update_on=['cellValueChanged'],
                          key=f'financial_grid_{fingerprint}')
        st.caption('Use Level 1 / 2 / 3 in the grid header. Click + or − beside a category to expand or collapse it. Click a name or amount to open its ledger.')
        returned = response.data
        if isinstance(returned, pd.DataFrame) and '_action' in returned.columns:
            actions = returned.loc[returned['_action'].fillna('').ne('')]
            if not actions.empty:
                action_row = actions.iloc[-1]
                action_path = paths[int(action_row['_id'])]
                action = str(action_row['_action'])
                token = (fingerprint, st.session_state.grid_revision, int(action_row['_id']), action)
                if action.startswith('ledger:') and st.session_state.get('handled_action') != token:
                    st.session_state.handled_action = token
                    st.session_state.ledger_path = action_path
                    opened = action_row.get('_expanded', '')
                    if opened:
                        st.session_state.expanded = {tuple(json.loads(item)) for item in json.loads(opened)}
                    clicked_column = action.split(':')[1]
                    if clicked_column in ('2026', '2025'):
                        st.session_state.ledger_year = int(clicked_column)
        path = st.session_state.ledger_path
        if path is None:
            st.info('Click a category name or amount to view its September ledger.')
            return
        st.divider()
        st.subheader('2. Ledger Details')
        st.write('Selected scope: ' + ' → '.join(path))
        year = st.radio('Select the ledger year', [2026, 2025], horizontal=True, key='ledger_year')
        mask = df[MAIN].eq(path[0]) & df['Year'].eq(year)
        if len(path) >= 2:
            mask &= df[SUB].eq(path[1])
        if len(path) == 3:
            mask &= df[CODE].eq(path[2])
        detail = df.loc[mask].copy().sort_values('Posting Date')
        a, b = st.columns(2)
        a.metric(f'Transactions · September {year}', len(detail))
        b.metric('Original Ledger Net Amount', f'{detail[AMOUNT].sum():,.2f}')
        detail['Debit'] = detail[AMOUNT].clip(lower=0)
        detail['Credit'] = -detail[AMOUNT].clip(upper=0)
        st.caption('Positive GL amounts appear in Debit; negative GL amounts appear as positive Credit values. Net = Debit − Credit.')
        columns = [CODE, 'Acct Name', MAIN, SUB, 'Title', 'Posting Date', 'Origin', 'Doc No',
                   'Remarks', 'Offset Acct', 'Offset Acct Name', 'Debit', 'Credit', AMOUNT]
        detail = detail[[col for col in columns if col in detail.columns]].copy()
        detail['Posting Date'] = detail['Posting Date'].dt.strftime('%Y-%m-%d')
        if detail.empty:
            st.info('No September transactions were found for this selection and year.')
        total_row = {col: '' for col in detail.columns}
        total_row[CODE] = 'Grand Total'
        for col in ['Debit', 'Credit', AMOUNT]:
            total_row[col] = float(detail[col].sum())
        detail = pd.concat([detail, pd.DataFrame([total_row])], ignore_index=True)
        st.dataframe(detail, use_container_width=True, hide_index=True,
                     column_config={col: st.column_config.NumberColumn(format='%.2f')
                                    for col in ['Debit', 'Credit', AMOUNT]})
        st.download_button('Download Selected Ledger (CSV)', detail.to_csv(index=False).encode('utf-8-sig'),
                           file_name=f'september_{year}_ledger.csv', mime='text/csv')
    except Exception as exc:
        st.error(f'Unable to process the workbook: {exc}')


if __name__ == '__main__':
    run()
