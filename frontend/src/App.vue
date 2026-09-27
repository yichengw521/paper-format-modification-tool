<script setup>
import { computed, onBeforeUnmount, ref } from 'vue'
import axios from 'axios'
import { ElMessage, ElMessageBox } from 'element-plus'

const templateFile = ref(null)
const documentFile = ref(null)
const task = ref(null)
const report = ref(null)
const plan = ref(null)
const submitting = ref(false)
const errorMessage = ref('')
let pollTimer = null

// 后端只返回状态码，前端负责转换成用户可读文案。
const statusText = computed(() => ({
  ANALYZING: '正在解析模板和文档',
  AWAITING_CONFIRMATION: '等待确认修改要求',
  QUEUED: '等待处理',
  PROCESSING: '正在分析并修改',
  COMPLETED: '处理完成',
  FAILED: '处理失败',
}[task.value?.status] || '准备上传'))

const activeStep = computed(() => {
  if (!task.value) return templateFile.value && documentFile.value ? 1 : 0
  if (task.value.status === 'AWAITING_CONFIRMATION') return 2
  if (task.value.status === 'COMPLETED') return 4
  if (task.value.status === 'QUEUED' || task.value.status === 'PROCESSING') return 3
  return 1
})

const canSubmit = computed(() => Boolean(templateFile.value && documentFile.value)
  && !submitting.value)

// Element Plus 这里没有使用 Upload 组件，保留原生 input 便于直接拿到 File 对象。
function chooseFile(event, kind) {
  const file = event.target.files?.[0]
  event.target.value = ''
  if (!file) return
  if (!file.name.toLowerCase().endsWith('.docx')) {
    ElMessage.error('请选择 DOCX 格式的 Word 文件')
    return
  }
  if (file.size > 50 * 1024 * 1024) {
    ElMessage.error('单个文件不能超过 50 MB')
    return
  }
  if (kind === 'template') templateFile.value = file
  else documentFile.value = file
  resetResult()
}

function resetResult() {
  task.value = null
  report.value = null
  plan.value = null
  errorMessage.value = ''
  submitting.value = false
  stopPolling()
}

function clearFile(kind) {
  if (kind === 'template') templateFile.value = null
  else documentFile.value = null
  resetResult()
}

function formatSize(bytes) {
  if (!bytes) return '0 KB'
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  return `${(bytes / 1024 / 1024).toFixed(2)} MB`
}

async function submitTask() {
  if (!canSubmit.value) return
  submitting.value = true
  errorMessage.value = ''
  report.value = null
  try {
    const form = new FormData()
    form.append('template', templateFile.value)
    form.append('document', documentFile.value)
    const response = await axios.post('/api/v1/tasks', form)
    task.value = response.data
    pollTask()
  } catch (error) {
    failWith(error)
  }
}

// 后端处理是异步任务；前端轮询 status 链接直到完成或失败。
async function pollTask() {
  stopPolling()
  if (!task.value?.links?.status) return
  try {
    const response = await axios.get(task.value.links.status)
    task.value = response.data
    if (task.value.status === 'AWAITING_CONFIRMATION') {
      submitting.value = false
      await loadPlan()
      ElMessage.success('模板规则已解析，请确认后再执行修改')
      return
    }
    if (task.value.status === 'COMPLETED') {
      submitting.value = false
      await loadReport()
      ElMessage.success('格式处理完成，可以下载新文档了')
      return
    }
    if (task.value.status === 'FAILED') {
      submitting.value = false
      errorMessage.value = task.value.error || '处理失败，请检查文件后重试。'
      return
    }
    pollTimer = window.setTimeout(pollTask, 900)
  } catch (error) {
    failWith(error)
  }
}

async function loadPlan() {
  const response = await axios.get(task.value.links.analysis)
  plan.value = response.data
}

async function confirmPlan() {
  const enabledRuleKeys = plan.value?.rules?.filter(rule => rule.enabled).map(rule => rule.key) || []
  const acceptedIssueKeys = plan.value?.issues?.filter(issue => issue.selected).map(issue => issue.key) || []
  if (!enabledRuleKeys.length) {
    ElMessage.warning('请至少选择一项需要修改的格式')
    return
  }
  try {
    await ElMessageBox.confirm(
      `确认按 ${enabledRuleKeys.length} 项格式规则处理，并采用 ${acceptedIssueKeys.length} 项章节文字修正吗？原文件不会被覆盖。`,
      '二次确认',
      { confirmButtonText: '确认并执行修改', cancelButtonText: '继续检查', type: 'warning' },
    )
    submitting.value = true
    const response = await axios.post(task.value.links.confirm, { enabledRuleKeys, acceptedIssueKeys })
    task.value = response.data
    pollTask()
  } catch (error) {
    if (error !== 'cancel' && error !== 'close') failWith(error)
  }
}

async function loadReport() {
  try {
    const response = await axios.get(task.value.links.report)
    report.value = response.data
  } catch {
    ElMessage.warning('文档已生成，但详细报告暂时无法加载')
  }
}

// 后端错误统一使用 ProblemDetail，用户可读错误在 detail 字段里。
function failWith(error) {
  submitting.value = false
  const detail = error?.response?.data?.detail
  errorMessage.value = detail || '无法连接处理服务，请确认后端已经启动。'
  stopPolling()
}

function stopPolling() {
  if (pollTimer) window.clearTimeout(pollTimer)
  pollTimer = null
}

// 结果和报告都是后端附件下载接口，使用新窗口避免打断当前结果页状态。
function download(url) {
  window.open(url, '_blank', 'noopener')
}

function startOver() {
  templateFile.value = null
  documentFile.value = null
  resetResult()
}

onBeforeUnmount(stopPolling)
</script>

<template>
  <div class="page-shell">
    <header class="topbar">
      <a class="brand" href="#">
        <span class="brand-mark">P</span>
        <span>
          <strong>Paper Format</strong>
          <small>Modification Tool</small>
        </span>
      </a>
    </header>
    <main>
      <section class="hero">
        <div class="eyebrow">论文格式自动修改工具</div>
        <h1>把格式交给工具，<br /><em>把时间留给内容。</em></h1>
        <p>上传学校模板和你的文档，系统会识别标题、正文与页面规则，生成一份新的 Word 文档。原文件始终保留。</p>
        <div class="feature-row">
          <span>✓ 不覆盖原文件</span>
          <span>✓ 保留图片与表格</span>
          <span>✓ 附带修改报告</span>
        </div>
      </section>

      <section class="workspace-card">
        <el-steps :active="activeStep" finish-status="success" align-center class="steps">
          <el-step title="选择文件" />
          <el-step title="解析模板" />
          <el-step title="确认要求" />
          <el-step title="自动处理" />
          <el-step title="下载结果" />
        </el-steps>

        <div v-if="!task || task.status === 'FAILED'" class="upload-area">
          <div class="upload-grid">
            <article class="file-card" :class="{ selected: templateFile }">
              <div class="file-icon template-icon">范</div>
              <div class="file-copy">
                <span class="file-kicker">第一步</span>
                <h2>学校格式模板</h2>
                <p>选择包含正确页面、标题和正文样式的 DOCX 文件。</p>
              </div>
              <div v-if="templateFile" class="selected-file">
                <div><strong>{{ templateFile.name }}</strong><small>{{ formatSize(templateFile.size) }}</small></div>
                <button type="button" aria-label="移除模板" @click="clearFile('template')">×</button>
              </div>
              <label v-else class="select-button">
                选择模板
                <input type="file" accept=".docx" @change="chooseFile($event, 'template')" />
              </label>
            </article>

            <article class="file-card" :class="{ selected: documentFile }">
              <div class="file-icon document-icon">文</div>
              <div class="file-copy">
                <span class="file-kicker">第二步</span>
                <h2>需要修改的说明书</h2>
                <p>选择待检查和修改的 DOCX，系统会另存为新文件。</p>
              </div>
              <div v-if="documentFile" class="selected-file">
                <div><strong>{{ documentFile.name }}</strong><small>{{ formatSize(documentFile.size) }}</small></div>
                <button type="button" aria-label="移除说明书" @click="clearFile('document')">×</button>
              </div>
              <label v-else class="select-button">
                选择说明书
                <input type="file" accept=".docx" @change="chooseFile($event, 'document')" />
              </label>
            </article>
          </div>

          <el-alert v-if="errorMessage" :title="errorMessage" type="error" show-icon :closable="false" />

          <div class="action-row">
            <p>支持 DOCX，单个文件不超过 50 MB</p>
            <el-button type="primary" size="large" :disabled="!canSubmit" :loading="submitting" @click="submitTask">
              解析模板规则
            </el-button>
          </div>
        </div>

        <div v-else-if="task.status === 'AWAITING_CONFIRMATION' && plan" class="confirmation-panel">
          <div class="confirmation-heading">
            <div>
              <span class="file-kicker">修改前确认</span>
              <h2>请确认本次采用的格式规则</h2>
              <p>系统现在还没有修改文档。请逐项检查模板解析结果，确认后才会生成新文件。</p>
            </div>
            <div class="document-counts">
              <span>{{ plan.documentSummary.bodyParagraphs }} 个正文段落</span>
              <span>{{ plan.documentSummary.tables }} 个表格</span>
              <span>{{ plan.documentSummary.sections }} 个分节</span>
            </div>
          </div>

          <div v-if="plan.issues?.length" class="issue-panel">
            <div class="issue-heading">
              <div>
                <span class="file-kicker">结构问题</span>
                <h3>请确认章节文字修正</h3>
              </div>
              <span>{{ plan.issues.length }} 项待确认</span>
            </div>
            <article v-for="issue in plan.issues" :key="issue.key" class="issue-card" :class="{ disabled: !issue.selected }">
              <el-checkbox v-model="issue.selected" size="large" />
              <div>
                <strong>{{ issue.originalText }} <em>→</em> {{ issue.suggestedText }}</strong>
                <p>{{ issue.reason }}</p>
                <small>{{ issue.confidence }}% 置信度 · 取消勾选将保留原文</small>
              </div>
            </article>
          </div>

          <div class="rule-list">
            <article v-for="rule in plan.rules" :key="rule.key" class="rule-card" :class="{ disabled: !rule.enabled }">
              <el-checkbox v-model="rule.enabled" size="large" />
              <div class="rule-copy">
                <div class="rule-title">
                  <span>{{ rule.group }}</span>
                  <strong>{{ rule.label }}</strong>
                  <i>{{ rule.confidence }}% 匹配</i>
                </div>
                <p>{{ rule.font }} · {{ rule.size }}</p>
                <small>{{ rule.paragraphFormatting }}</small>
                <details><summary>查看识别依据</summary><p>{{ rule.evidence }}</p></details>
              </div>
            </article>
          </div>

          <el-alert title="目录会保留为 Word 自动目录；本机安装 Microsoft Word 时会自动刷新页码和点引导符。" type="info" show-icon :closable="false" />
          <div class="confirm-actions">
            <button class="text-button" type="button" @click="startOver">重新选择文件</button>
            <el-button type="primary" size="large" :loading="submitting" @click="confirmPlan">确认并执行修改</el-button>
          </div>
        </div>

        <div v-else-if="task.status !== 'COMPLETED'" class="processing-panel">
          <div class="orbit"><span></span><i></i></div>
          <span class="file-kicker">{{ statusText }}</span>
          <h2>{{ task.status === 'ANALYZING' ? '正在读取模板规则' : '正在生成新文档' }}</h2>
          <p>{{ task.status === 'ANALYZING' ? '系统只做结构和格式分析，分析完成后会请你确认。' : '系统正在按已确认的规则修改封面、摘要、目录、正文和后置部分。' }}</p>
          <el-progress :percentage="task.status === 'ANALYZING' ? 42 : (task.status === 'QUEUED' ? 58 : 82)" :show-text="false" :stroke-width="8" />
          <small>任务编号 {{ task.id }}</small>
        </div>

        <div v-else class="result-panel">
          <div class="success-heading">
            <div class="success-icon">✓</div>
            <div><span class="file-kicker">处理完成</span><h2>新文档已经生成</h2><p>原始模板和文档没有被覆盖。</p></div>
          </div>

          <div class="summary-grid">
            <div><strong>{{ task.modifications ?? 0 }}</strong><span>自动修正</span></div>
            <div><strong>{{ task.tables ?? 0 }}</strong><span>检查表格</span></div>
            <div><strong>{{ task.sections ?? 0 }}</strong><span>检查分节</span></div>
            <div><strong>{{ report?.documentSummary?.images ?? '—' }}</strong><span>保留图片</span></div>
          </div>

          <div class="download-row">
            <el-button type="primary" size="large" @click="download(task.links.result)">下载修改后的文档</el-button>
            <el-button size="large" @click="download(task.links.report)">下载修改报告</el-button>
            <button class="text-button" type="button" @click="startOver">处理另一份文件</button>
          </div>

          <div v-if="report" class="report-preview">
            <div class="report-title"><div><span class="file-kicker">本次修改</span><h3>格式调整明细</h3></div><span>{{ report.modifications.length }} 项</span></div>
            <div v-if="report.modifications.length" class="change-list">
              <div v-for="item in report.modifications" :key="item.bodyParagraphIndex" class="change-item">
                <span class="change-index">{{ String(item.bodyParagraphIndex).padStart(3, '0') }}</span>
                <div><strong>{{ item.textPreview }}</strong><p>{{ item.beforeStyle || '无样式' }} → {{ item.afterStyle }}</p></div>
              </div>
            </div>
            <p v-else class="empty-copy">没有发现需要自动修改的明确格式问题。</p>

            <div v-if="report.manualReviewItems?.length" class="review-box">
              <h3>建议人工确认</h3>
              <ul><li v-for="item in report.manualReviewItems" :key="item">{{ item }}</li></ul>
            </div>
          </div>
        </div>
      </section>

      <section class="how-it-works">
        <div><span>01</span><h3>读取模板</h3><p>提取页面、字体、字号、行距、缩进与标题规则。</p></div>
        <div><span>02</span><h3>识别结构</h3><p>区分摘要、标题、正文、表题和参考文献等内容。</p></div>
        <div><span>03</span><h3>安全修改</h3><p>只修正确认度高的项目，并输出一份独立的新文档。</p></div>
      </section>
    </main>

    <footer>Paper Format Modification Tool · 当前为本机 MVP 版本</footer>
  </div>
</template>
