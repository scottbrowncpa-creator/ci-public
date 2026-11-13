import { expect, test } from 'claude-code/testing'

// Colors, width and text styles all come from hooks/settings.ts.
// One sample reply on a 200-column window: heading, bold, italic, a link, a rule, a table, a mermaid block.
const REPLY = [
  '## Cash summary',
  '',
  'Fund 50 balances to **zero** this month, *slow* week, see [docs](https://example.com).',
  '',
  '| Fund | Status | Note |',
  '| --- | --- | --- |',
  '| 01 | ties | ' + 'money in money out what we owe '.repeat(8) + '|',
  '| 50 | ties | short |',
  '',
  '```mermaid',
  'flowchart LR',
  '  A[Bank] --> B[GL]',
  '```',
].join('\n')

const log = (globalThis as any).console.log as (s: string) => void

test('a sample reply follows settings.ts: colors, width and text styles', async ($, on) => {
  on('session.start', (_: any, e: any) => ({ cwd: e.cwd }))
  on('env.get', () => ({ value: undefined }))
  await $.session.start({ cwd: '/', surface: 'terminal', isInteractive: true })
  const ui = await $.ui.mount({
    plugin: 'ci-colors', component: 'AssistantMessage', surface: 'terminal',
    props: { text: REPLY, isFirstOfReply: true },
    viewport: { columns: 200, rows: 60 },
  } as any)
  const texts = (await ui.findAll({ type: 'Text' })) as any[]
  const named = (re: RegExp) => texts.filter(t => re.test(t.text ?? ''))
  const colorFor = (re: RegExp) => named(re).map(t => String(t.props.color ?? '').toLowerCase())

  // colors
  expect(colorFor(/Cash summary/)).toContain('#5bb0c4')
  expect(colorFor(/^zero$/)).toContain('#f5c449')
  expect(colorFor(/^Status$/)).toContain('#f5c449')
  expect(colorFor(/Bank|┌/).some(c => c === '#5bb0c4')).toBe(true)

  // text styles
  expect(named(/^zero$/).some(t => t.props.bold === true)).toBe(true)
  expect(named(/^slow$/).some(t => t.props.italic === true)).toBe(true)
  expect(named(/^docs$/).some(t => t.props.underline === true)).toBe(true)
  expect(named(/^━+$/).length).toBeGreaterThan(0) // banner heading style: heavy rule under H2

  // width: the window is 200 columns, the cap is 120
  const widest = Math.max(...texts.map(t => String(t.text ?? '').length))
  const rule = Math.max(...named(/^[┌└╞├][─═┬┼╪┴]+[┐┘╡┤]$/).map(t => t.text.length), 0)
  const box = (await ui.findAll({ type: 'Box', key: 'ciWidth' } as any))[0] as any
  log(`widest drawn line ${widest} columns, wide table ${rule} columns, reply box width ${box?.props.width}, window 200`)
  expect(box?.props.width).toBeLessThanOrEqual(120)
  log(texts.filter(t => /zero|slow|docs|Cash|Status/.test(t.text ?? '')).map(t => `${t.text} color=${t.props.color} bold=${!!t.props.bold} italic=${!!t.props.italic} underline=${!!t.props.underline}`).join('\n'))
  expect(widest).toBeLessThanOrEqual(120)
  expect(rule).toBeGreaterThan(100)
  await ui.unmount()
})
