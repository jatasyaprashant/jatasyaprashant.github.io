(() => {
  const byId = id => document.getElementById(id);
  const element = (tag, text) => {
    const node = document.createElement(tag);
    if (text !== undefined) node.textContent = text;
    return node;
  };

  // Use the same displayed percentages and counts as the verified static tables.
  const populations = new Map();
  const windowSelect = byId('window');
  for (const section of document.querySelectorAll('details[data-window]')) {
    const table = section.querySelector('table');
    const years = Array.from(table.tHead.rows[0].cells).slice(1).map(cell => cell.textContent);
    const [totalRow, ...rows] = table.tBodies[0].rows;
    const totals = Array.from(totalRow.cells).slice(1).map(cell => Number(cell.textContent));
    const domains = new Map();
    for (const row of rows) {
      domains.set(row.cells[0].textContent, Array.from(row.cells).slice(1).map(cell => ({
        count: Number(cell.firstChild.textContent),
        share: cell.querySelector('small').textContent
      })));
    }
    populations.set(section.dataset.window, {years, totals, domains});
    windowSelect.add(new Option(section.querySelector('summary').textContent, section.dataset.window));
  }
  const first = byId('domain-a');
  const second = byId('domain-b');
  for (const name of populations.get('matched').domains.keys()) {
    first.add(new Option(name));
    second.add(new Option(name));
  }
  first.value = 'NLP/LLM';
  second.value = 'Agents';
  let yearIndex = populations.get('matched').years.length - 1;

  function updateComparison() {
    const {years, totals, domains} = populations.get(windowSelect.value);
    const names = [...new Set([first.value, second.value])];
    const measure = byId('measure').value;
    const number = entry => measure === 'count' ? entry.count : parseFloat(entry.share);
    const peak = Math.max(...names.flatMap(name => domains.get(name).map(number)));
    const scale = Math.max(5, Math.ceil(peak / 5) * 5);
    byId('axis').textContent = windowSelect.selectedOptions[0].textContent +
      ' · Shared scale from 0 to ' + scale + (measure === 'count' ? ' works' : '%');
    const blocks = names.map((name, index) => {
      const block = element('div');
      block.append(element('h4', name));
      domains.get(name).forEach((entry, i) => {
        const row = element('div');
        row.className = index ? 'bar-row all' : 'bar-row';
        const track = element('span');
        track.className = 'bar-track';
        track.setAttribute('aria-hidden', 'true');
        const bar = element('span');
        bar.style.setProperty('--size', (100 * number(entry) / scale) + '%');
        track.append(bar);
        row.append(element('span', years[i]), track, element('strong', entry[measure]));
        block.append(row);
      });
      return block;
    });
    byId('comparison').replaceChildren(...blocks);
    const readings = names.map(name => {
      const point = domains.get(name)[yearIndex];
      return name + ' accounted for ' + point.count + ' of ' + totals[yearIndex] +
        ' eligible works (' + point.share + ').';
    });
    byId('inspection').textContent = years[yearIndex] + ' · ' + readings.join(' ');
    Array.from(byId('years').children).forEach((button, i) =>
      button.setAttribute('aria-pressed', String(i === yearIndex)));
  }
  populations.get('matched').years.forEach((year, i) => {
    const button = element('button', year);
    button.type = 'button';
    button.addEventListener('click', () => { yearIndex = i; updateComparison(); });
    byId('years').append(button);
  });
  byId('explorer').addEventListener('change', updateComparison);
  updateComparison();
  byId('explorer').hidden = false;

  const groups = [...document.querySelectorAll('.focus-area')];
  const papers = [...document.querySelectorAll('.papers li')];
  for (const [id, key] of [['paper-year', 'year'], ['paper-status', 'status']]) {
    const values = [...new Set(papers.map(paper => paper.dataset[key]))].sort();
    if (key === 'year') values.reverse();
    values.forEach(value => byId(id).add(new Option(value)));
  }
  function filterPapers() {
    const query = byId('paper-search').value.trim().toLowerCase();
    const year = byId('paper-year').value;
    const status = byId('paper-status').value;
    const filtering = Boolean(query || year || status);
    const unique = new Set();
    for (const group of groups) {
      let visible = 0;
      for (const paper of group.querySelectorAll('li')) {
        paper.hidden = Boolean((year && paper.dataset.year !== year) ||
          (status && paper.dataset.status !== status) || !paper.textContent.toLowerCase().includes(query));
        if (!paper.hidden) { visible++; unique.add(paper.querySelector('a').href); }
      }
      group.hidden = visible === 0;
      group.querySelector('summary').textContent = 'Browse ' + visible + (visible === 1 ? ' paper' : ' papers');
      group.querySelector('details').open = filtering && visible > 0;
    }
    byId('paper-results').textContent = unique.size + ' distinct paper links across the focus lists.';
  }
  byId('paper-search').addEventListener('input', filterPapers);
  byId('paper-filters').addEventListener('change', filterPapers);
  byId('clear-papers').addEventListener('click', () => {
    ['paper-search','paper-year','paper-status'].forEach(id => { byId(id).value = ''; });
    filterPapers();
  });
  filterPapers();
  byId('paper-filters').hidden = false;

  const directory = byId('people-filter').closest('section');
  byId('people-search').addEventListener('input', event => {
    const query = event.target.value.trim().toLowerCase();
    let hits = 0;
    for (const group of directory.querySelectorAll('details')) {
      const city = group.querySelector('summary').firstChild.textContent.toLowerCase();
      let visible = 0;
      for (const row of group.querySelectorAll('tbody tr')) {
        const fields = [city, ...Array.from(row.cells, cell => cell.textContent.toLowerCase())];
        row.hidden = !fields.some(field => field.includes(query));
        if (!row.hidden) visible++;
      }
      hits += visible;
      group.hidden = visible === 0;
      group.open = Boolean(query && visible);
    }
    byId('people-results').textContent = query ? hits + (hits === 1 ? ' matching researcher.' : ' matching researchers.') : '';
  });
  byId('people-filter').hidden = false;
})();
