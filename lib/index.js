import { spawn } from 'node:child_process'
import { existsSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

export const name = 'dsh-pptx-anim'

/**
 * Host half of dsh-pptx-anim.
 *
 * Why this exists: python-pptx (and therefore most DSH pptx tooling) cannot write
 * animations — <p:timing> is not implemented. This plugin shells out to a small,
 * verified Python helper that injects the OOXML timing tree, plus a structural
 * validator that catches the mistakes PowerPoint would report as "needs repair".
 */

const HERE = dirname(fileURLToPath(import.meta.url))
const SCRIPTS = join(HERE, '..', 'scripts')
const ANIMATE = join(SCRIPTS, 'pptx_anim.py')
const VALIDATE = join(SCRIPTS, 'validate.py')

const PRESETS = ['auto', 'title', 'proposition', 'image', 'cards', 'data', 'none']

function candidates(override) {
  const list = []
  if (override) list.push(override)
  if (process.env.DSH_PPTX_ANIM_PYTHON) list.push(process.env.DSH_PPTX_ANIM_PYTHON)
  list.push('python', 'python3', 'py')
  return list
}

function run(exe, args, timeoutMs = 180000) {
  return new Promise((resolve) => {
    let stdout = ''
    let stderr = ''
    let settled = false
    let child
    try {
      child = spawn(exe, args, { windowsHide: true })
    } catch (error) {
      resolve({ ok: false, code: null, stdout: '', stderr: String(error && error.message) })
      return
    }
    const finish = (result) => {
      if (!settled) {
        settled = true
        resolve(result)
      }
    }
    const timer = setTimeout(() => {
      try {
        child.kill()
      } catch {
        /* ignore */
      }
      finish({ ok: false, code: null, stdout, stderr: `${stderr}\n[timed out after ${timeoutMs}ms]` })
    }, timeoutMs)

    child.stdout?.on('data', (chunk) => {
      stdout += String(chunk)
    })
    child.stderr?.on('data', (chunk) => {
      stderr += String(chunk)
    })
    child.on('error', (error) => {
      clearTimeout(timer)
      finish({ ok: false, code: null, stdout, stderr: `${stderr}\n${error && error.message}` })
    })
    child.on('close', (code) => {
      clearTimeout(timer)
      finish({ ok: code === 0, code, stdout, stderr })
    })
  })
}

/**
 * Try each interpreter until one both runs and can import python-pptx.
 * Returns { exe, version } or { error, tried }.
 */
async function pickPython(override) {
  const tried = []
  for (const exe of candidates(override)) {
    const probe = await run(exe, ['-c', 'import sys, pptx; print(sys.version.split()[0])'], 20000)
    const version = (probe.stdout || '').trim().split('\n')[0]
    tried.push({ exe, ok: probe.ok, version: version || undefined })
    if (probe.ok) return { exe, version, tried }
  }
  return { error: 'no interpreter with python-pptx found', tried }
}

function textResult(lines) {
  return [{ type: 'text', text: lines.filter(Boolean).join('\n') }]
}

async function animateHandler(args) {
  if (!existsSync(ANIMATE)) throw new Error(`helper script missing: ${ANIMATE}`)
  const input = String(args?.input || '').trim()
  if (!input) throw new Error('input is required (absolute path to a .pptx)')
  if (!existsSync(input)) throw new Error(`input not found: ${input}`)

  const output = String(args?.output || '').trim() || input.replace(/\.pptx$/i, '') + '_anim.pptx'
  const preset = PRESETS.includes(args?.preset) ? args.preset : 'auto'
  const transition = args?.transition === 'none' ? 'none' : 'fade'

  const picked = await pickPython(args?.python)
  if (picked.error) {
    return textResult([
      '没有找到可用的 Python（且需能 import pptx）。',
      '已尝试：' + picked.tried.map((t) => `${t.exe}${t.ok ? '' : '(失败)'}`).join(', '),
      '请在调用时传 python 参数（例如 DSH 自带运行时里的 python.exe），或设置环境变量 DSH_PPTX_ANIM_PYTHON。',
    ])
  }

  const runArgs = [ANIMATE, input, output, `--preset=${preset}`, `--transition=${transition}`, '--audit']
  const result = await run(picked.exe, runArgs)

  if (!result.ok) {
    return textResult([
      `注入失败（${picked.exe}，退出码 ${result.code}）。`,
      result.stdout ? '--- stdout ---\n' + result.stdout.trim() : '',
      result.stderr ? '--- stderr ---\n' + result.stderr.trim() : '',
    ])
  }

  return textResult([
    `已注入动画：${output}`,
    `Python: ${picked.exe} ${picked.version || ''}`.trim(),
    `preset=${preset}  transition=${transition}`,
    '',
    result.stdout.trim(),
    '',
    '提醒：本工具无法代验 PowerPoint 实机播放，首次放映前请按 F5 预览；原始文件未被修改。',
  ])
}

async function validateHandler(args) {
  if (!existsSync(VALIDATE)) throw new Error(`validator missing: ${VALIDATE}`)
  const input = String(args?.input || '').trim()
  if (!input) throw new Error('input is required (absolute path to a .pptx)')
  if (!existsSync(input)) throw new Error(`input not found: ${input}`)

  const picked = await pickPython(args?.python)
  if (picked.error) {
    return textResult([
      '没有找到可用的 Python（且需能 import pptx）。',
      '请在调用时传 python 参数，或设置 DSH_PPTX_ANIM_PYTHON。',
    ])
  }

  const result = await run(picked.exe, [VALIDATE, input])
  return textResult([
    result.ok ? '结构校验通过。' : '结构校验发现问题。',
    '',
    (result.stdout || '').trim(),
    result.stderr ? '--- stderr ---\n' + result.stderr.trim() : '',
  ])
}

async function envHandler(args) {
  const picked = await pickPython(args?.python)
  if (picked.error) {
    return textResult([
      '未找到可用 Python：',
      ...picked.tried.map((t) => `  ${t.exe} -> ${t.ok ? 'ok ' + (t.version || '') : '失败'}`),
    ])
  }
  return textResult([
    '可用 Python：',
    `  ${picked.exe}  (${picked.version || '?'})`,
    '',
    '探测记录：',
    ...picked.tried.map((t) => `  ${t.exe} -> ${t.ok ? 'ok ' + (t.version || '') : '失败'}`),
  ])
}

function registerTools(ctx) {
  ctx.inject(['tools'], (scope) => {
    const tools = [
      {
        name: 'pptx_animate',
        description:
          '给 .pptx 批量注入真正的 PowerPoint 进入动画与幻灯片切换。' +
          'python-pptx 没有动画 API，本工具直接写 OOXML 的 p:timing/p:transition 树。' +
          'input 为源文件绝对路径，output 省略则生成 <原名>_anim.pptx（不修改源文件）。' +
          'preset 可选 auto（按每页构成自动判定）/ title（封面）/ proposition（命题大字页）/' +
          ' image（图文页）/ cards（三栏卡片页）/ data（数据密集页）/ none（不注入）。' +
          'transition 默认 fade。返回每页动画节点数，调用后建议再用 pptx_anim_validate 校验。',
        parameters: {
          type: 'object',
          additionalProperties: false,
          properties: {
            input: { type: 'string', description: '源 .pptx 的绝对路径' },
            output: { type: 'string', description: '输出路径，省略则 <原名>_anim.pptx' },
            preset: { type: 'string', enum: PRESETS, description: '动画预设，默认 auto' },
            transition: { type: 'string', enum: ['fade', 'none'], description: '页面切换，默认 fade' },
            python: { type: 'string', description: '可选：指定 Python 解释器绝对路径（需可 import pptx）' },
          },
          required: ['input'],
        },
        output: {
          schema: { type: 'object', additionalProperties: true, required: ['ok'], properties: { ok: { type: 'boolean' } } },
          render: (_args, value) => textResult([String(value?.message || '')]),
        },
        async execute(args) {
          const lines = await animateHandler(args)
          return { ok: !String(lines[0]?.text || '').includes('失败'), message: lines[0]?.text || '' }
        },
        presentCall() {
          return { card: 'generic', title: 'pptx_animate', kind: 'read' }
        },
      },
      {
        name: 'pptx_anim_validate',
        description:
          '校验一个（主张已注入动画的）.pptx 结构：包能否重开、CT_Slide 子元素顺序（transition 必须在 timing 之前）、' +
          '空 stCondLst、每个动画 spTgt 是否指向真实形状、每页 animEffect 节点数。' +
          '这些问题一旦存在，PowerPoint 打开时会提示"需要修复"。注入后必跑。',
        parameters: {
          type: 'object',
          additionalProperties: false,
          properties: {
            input: { type: 'string', description: '.pptx 绝对路径' },
            python: { type: 'string', description: '可选：Python 解释器绝对路径' },
          },
          required: ['input'],
        },
        output: {
          schema: { type: 'object', additionalProperties: true, required: ['ok'], properties: { ok: { type: 'boolean' } } },
          render: (_args, value) => textResult([String(value?.message || '')]),
        },
        async execute(args) {
          const lines = await validateHandler(args)
          return { ok: String(lines[0]?.text || '').includes('通过'), message: lines[0]?.text || '' }
        },
        presentCall() {
          return { card: 'generic', title: 'pptx_anim_validate', kind: 'read' }
        },
      },
      {
        name: 'pptx_anim_env',
        description:
          '探测哪个 Python 解释器可用（且能 import python-pptx），用于 pptx_animate / pptx_anim_validate 的 python 参数。',
        parameters: {
          type: 'object',
          additionalProperties: false,
          properties: {
            python: { type: 'string', description: '可选：要验证的解释器绝对路径' },
          },
        },
        output: {
          schema: { type: 'object', additionalProperties: true, required: ['ok'], properties: { ok: { type: 'boolean' } } },
          render: (_args, value) => textResult([String(value?.message || '')]),
        },
        async execute(args) {
          const lines = await envHandler(args)
          return { ok: String(lines[0]?.text || '').includes('可用 Python'), message: lines[0]?.text || '' }
        },
        presentCall() {
          return { card: 'generic', title: 'pptx_anim_env', kind: 'read' }
        },
      },
    ]

    for (const tool of tools) {
      try {
        scope.tools.register(tool)
      } catch (error) {
        console.error('[dsh-pptx-anim] register failed for', tool.name, error)
      }
    }
    console.log('[dsh-pptx-anim] tools registered:', tools.length)
  })
}

export function apply(ctx) {
  try {
    registerTools(ctx)
  } catch (error) {
    console.error('[dsh-pptx-anim] apply failed', error)
  }
}
