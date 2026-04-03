import { LitElement, css, html } from 'lit'
import { customElement, property, state } from 'lit/decorators.js'

type QueueTask = {
  id: string
  status: string
  origin?: string
  destination?: string
  createdTime?: number | string
  contactPriority?: number
  transferCount?: number
  connectedCount?: number
  connectedDuration?: number
  owner?: {
    id?: string
    name?: string
  }
  lastQueue?: { name?: string }
  customer?: { name?: string; phoneNumber?: string }
  requiredSkills?: Array<{
    operand?: string
    name?: string
    intVal?: number
  }>
  callbackData?: { callbackStatus?: string }
  waitTimeMs?: number
  waitTimeMinutes?: number
}

type ActiveAgentSession = {
  agentName?: string
  teamName?: string
  agentSkills?: unknown
  startTime?: number | string
  channelInfo?: Array<{
    currentState?: string
    idleCodeName?: string
    lastActivityTime?: number | string
  }>
}

type CallSortColumn =
  | 'wait'
  | 'queue'
  | 'ani'
  | 'customer'
  | 'status'
  | 'priority'
  | 'transfers'
  | 'connected'
  | 'connectedDuration'
  | 'requiredSkills'
  | 'callback'
  | 'taskId'

type AgentSortColumn =
  | 'agentName'
  | 'teamName'
  | 'currentState'

type AgentStateFilter = 'all' | 'available' | 'idle' | 'rona' | 'wrapup'

const DEFAULT_SEARCH_URL = 'https://api.wxcc-us1.cisco.com/search'
const DEFAULT_STATUS = 'parked'
const DEFAULT_LOOKBACK_MINUTES = 24 * 60
const DEFAULT_REFRESH_MS = 30000

const GRAPHQL_QUERY_PARKED = `
query TaskDetailsParkedOverThreshold(
  $from: Long!
  $to: Long!
  $filter: TaskDetailsFilters
) {
  taskDetails(
    from: $from
    to: $to
    timeComparator: createdTime
    filter: $filter
  ) {
    tasks {
      id
      status
      isActive
      origin
      destination
      createdTime
      contactPriority
      transferCount
      connectedCount
      connectedDuration
      owner {
        id
        name
      }
      lastQueue {
        name
      }
      customer {
        name
        phoneNumber
      }
      requiredSkills
      callbackData {
        callbackStatus
      }
    }
  }
}
`

const GRAPHQL_QUERY_ACTIVE_AGENTS = `
query activeAgents(
  $from: Long!
  $to: Long!
  $filter: AgentSessionFilters
  $extFilter: AgentSessionSpecificFilters
  $pagination: Pagination
  ) {
    agentSession(
      from: $from
    to: $to
    filter: $filter
    extFilter: $extFilter
    pagination: $pagination
  ) {
    agentSessions {
      agentSkills
      agentName
      teamName
      startTime
      channelInfo {
        currentState
        idleCodeName
        lastActivityTime
      }
    }
  }
}
`

@customElement('queue-threshold-dashboard')
export class MyElement extends LitElement {
  @property() token = ''
  @property({ attribute: 'status-filter' }) statusFilter = DEFAULT_STATUS
  @property({ attribute: 'search-url' }) searchUrl = DEFAULT_SEARCH_URL
  @property({ attribute: 'lookback-minutes', type: Number }) lookbackMinutes =
    DEFAULT_LOOKBACK_MINUTES
  @property({ attribute: 'refresh-ms', type: Number }) refreshMs =
    DEFAULT_REFRESH_MS
  @property({ attribute: 'darkmode' }) darkmode = 'false'

  @state() private loading = false
  @state() private error = ''
  @state() private tasks: QueueTask[] = []
  @state() private connectedTasks: QueueTask[] = []
  @state() private activeAgents: ActiveAgentSession[] = []
  @state() private lastUpdated = ''
  @state() private apiRequestCount = 0
  @state() private callSortColumn: CallSortColumn = 'wait'
  @state() private callSortDirection: 'asc' | 'desc' = 'desc'
  @state() private connectedCallSortColumn: CallSortColumn = 'wait'
  @state() private connectedCallSortDirection: 'asc' | 'desc' = 'desc'
  @state() private agentSortColumn: AgentSortColumn = 'agentName'
  @state() private agentSortDirection: 'asc' | 'desc' = 'asc'
  @state() private agentStateFilter: AgentStateFilter = 'all'
  @state() private agentTeamFilter = 'all'
  @state() private agentIdleCodeFilter = 'all'

  private timerId?: number

  connectedCallback() {
    super.connectedCallback()
    this.loadData()
    this.startRefreshTimer()
  }

  disconnectedCallback() {
    super.disconnectedCallback()
    if (this.timerId) {
      window.clearInterval(this.timerId)
    }
  }

  updated(changedProperties: Map<string, unknown>) {
    if (
      changedProperties.has('refreshMs') ||
      changedProperties.has('token') ||
      changedProperties.has('statusFilter') ||
      changedProperties.has('searchUrl') ||
      changedProperties.has('lookbackMinutes')
    ) {
      this.startRefreshTimer()
      this.loadData()
    }
  }

  private startRefreshTimer() {
    if (this.timerId) {
      window.clearInterval(this.timerId)
    }

    if (!Number.isFinite(this.refreshMs) || this.refreshMs <= 0) {
      return
    }

    this.timerId = window.setInterval(() => this.loadData(), this.refreshMs)
  }

  private async loadData() {
    if (!this.token) {
      this.error = 'Missing access token.'
      this.tasks = []
      this.connectedTasks = []
      this.activeAgents = []
      return
    }

    this.loading = true
    this.error = ''

    const nowMs = Date.now()
    const queueToMs = nowMs
    const queueFromMs = queueToMs - this.lookbackMinutes * 60 * 1000

    const queuePayload = {
      query: GRAPHQL_QUERY_PARKED,
      variables: {
        from: queueFromMs,
        to: queueToMs,
        filter: {
          and: [
            { isActive: { equals: true } },
            { status: { equals: this.statusFilter } },
          ],
        },
      },
    }

    const connectedCallsPayload = {
      query: GRAPHQL_QUERY_PARKED,
      variables: {
        from: queueFromMs,
        to: queueToMs,
        filter: {
          and: [
            { isActive: { equals: true } },
            { status: { equals: 'connected' } },
          ],
        },
      },
    }

    const activeAgentsPayload = {
      query: GRAPHQL_QUERY_ACTIVE_AGENTS,
      variables: {
        from: nowMs - 24 * 60 * 60 * 1000,
        to: nowMs,
        filter: {
          and: [
            { isActive: { equals: true } },
            { channelInfo: { channelType: { equals: 'telephony' } } },
          ],
        },
      },
    }

    try {
      this.apiRequestCount += 3
      const [queueResponse, connectedCallsResponse, activeAgentsResponse] = await Promise.all([
        fetch(this.searchUrl, {
          method: 'POST',
          headers: {
            Authorization: `Bearer ${this.token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(queuePayload),
        }),
        fetch(this.searchUrl, {
          method: 'POST',
          headers: {
            Authorization: `Bearer ${this.token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(connectedCallsPayload),
        }),
        fetch(this.searchUrl, {
          method: 'POST',
          headers: {
            Authorization: `Bearer ${this.token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(activeAgentsPayload),
        }),
      ])

      if (!queueResponse.ok) {
        const queueErrorText = await queueResponse.text()
        throw new Error(
          `Queue search failed with status ${queueResponse.status}: ${queueErrorText || 'No response body'}`
        )
      }

      if (!activeAgentsResponse.ok) {
        const activeAgentsErrorText = await activeAgentsResponse.text()
        throw new Error(
          `Active agent search failed with status ${activeAgentsResponse.status}: ${activeAgentsErrorText || 'No response body'}`
        )
      }

      if (!connectedCallsResponse.ok) {
        const connectedCallsErrorText = await connectedCallsResponse.text()
        throw new Error(
          `Connected calls search failed with status ${connectedCallsResponse.status}: ${connectedCallsErrorText || 'No response body'}`
        )
      }

      const queueResult = (await queueResponse.json()) as {
        errors?: Array<{ message?: string }>
        data?: { taskDetails?: { tasks?: QueueTask[] } }
      }
      const connectedCallsResult = (await connectedCallsResponse.json()) as {
        errors?: Array<{ message?: string }>
        data?: { taskDetails?: { tasks?: QueueTask[] } }
      }
      const activeAgentsResult = (await activeAgentsResponse.json()) as {
        errors?: Array<{ message?: string }>
        data?: { agentSession?: { agentSessions?: ActiveAgentSession[] } }
      }

      if (queueResult.errors?.length) {
        throw new Error(queueResult.errors.map((item) => item.message).join(', '))
      }

      if (activeAgentsResult.errors?.length) {
        throw new Error(activeAgentsResult.errors.map((item) => item.message).join(', '))
      }

      if (connectedCallsResult.errors?.length) {
        throw new Error(connectedCallsResult.errors.map((item) => item.message).join(', '))
      }

      const rawTasks = queueResult.data?.taskDetails?.tasks ?? []
      const tasks = rawTasks
        .map((task) => this.addWaitFields(task, nowMs))
        .sort((left, right) => (right.waitTimeMs ?? 0) - (left.waitTimeMs ?? 0))

      const rawConnectedTasks = connectedCallsResult.data?.taskDetails?.tasks ?? []
      const connectedTasks = rawConnectedTasks
        .map((task) => this.addWaitFields(task, nowMs))
        .sort((left, right) => (right.waitTimeMs ?? 0) - (left.waitTimeMs ?? 0))

      const activeAgents =
        activeAgentsResult.data?.agentSession?.agentSessions
          ?.slice()
          .sort((left, right) =>
            String(left.agentName ?? '').localeCompare(String(right.agentName ?? ''))
          ) ?? []

      this.tasks = tasks
      this.connectedTasks = connectedTasks
      this.activeAgents = activeAgents
      this.lastUpdated = new Date(nowMs).toLocaleTimeString()
    } catch (error) {
      this.tasks = []
      this.connectedTasks = []
      this.activeAgents = []
      this.error =
        error instanceof Error ? error.message : 'Unable to load dashboard data.'
    } finally {
      this.loading = false
    }
  }

  private addWaitFields(task: QueueTask, nowMs: number): QueueTask {
    const createdMs = Number(task.createdTime ?? nowMs)
    const waitTimeMs = Number.isFinite(createdMs) ? Math.max(0, nowMs - createdMs) : 0
    const waitTimeMinutes = Math.floor(waitTimeMs / 60000)

    return {
      ...task,
      waitTimeMs,
      waitTimeMinutes,
    }
  }

  private formatWait(waitTimeMs?: number) {
    if (!waitTimeMs && waitTimeMs !== 0) {
      return 'Unknown'
    }

    const minutes = Math.floor(waitTimeMs / 60000)
    return String(minutes)
  }

  private formatConnectedDuration(durationMs?: number) {
    if (!durationMs) {
      return '0m 00s'
    }

    return this.formatWait(durationMs)
  }

  private formatRequiredSkills(skills?: QueueTask['requiredSkills']) {
    if (!Array.isArray(skills) || skills.length === 0) {
      return 'None'
    }

    const descriptions = skills
      .map((skill) => {
        const name = skill.name?.trim()
        if (!name) {
          return ''
        }

        const operand = skill.operand?.trim()
        const value =
          skill.intVal !== undefined && skill.intVal !== null ? String(skill.intVal) : ''

        if (operand && value) {
          return `${name} (${operand} ${value})`
        }

        if (operand) {
          return `${name} (${operand})`
        }

        return name
      })
      .filter((description): description is string => Boolean(description))

    return descriptions.length > 0 ? descriptions.join(', ') : 'None'
  }

  private getPrimaryChannel(agent: ActiveAgentSession) {
    return Array.isArray(agent.channelInfo) ? agent.channelInfo[0] : undefined
  }

  private async copyText(value: string) {
    try {
      await navigator.clipboard.writeText(value)
    } catch {
      console.warn('Clipboard copy failed')
    }
  }

  private getNormalizedAgentState(agent: ActiveAgentSession) {
    return String(this.getPrimaryChannel(agent)?.currentState ?? '')
      .trim()
      .toLowerCase()
  }

  private getAgentStatePriority(agent: ActiveAgentSession) {
    const state = this.getNormalizedAgentState(agent)
    switch (state) {
      case 'available':
        return 0
      case 'rona':
        return 1
      case 'idle':
        return 2
      case 'wrapup':
        return 3
      default:
        return 9
    }
  }

  private setAgentStateFilter(event: Event) {
    const target = event.target as HTMLSelectElement
    this.agentStateFilter = target.value as AgentStateFilter
  }

  private setAgentTeamFilter(event: Event) {
    const target = event.target as HTMLSelectElement
    this.agentTeamFilter = target.value
  }

  private setAgentIdleCodeFilter(event: Event) {
    const target = event.target as HTMLSelectElement
    this.agentIdleCodeFilter = target.value
  }

  private toggleCallSort(column: CallSortColumn) {
    if (this.callSortColumn === column) {
      this.callSortDirection = this.callSortDirection === 'asc' ? 'desc' : 'asc'
      return
    }

    this.callSortColumn = column
    this.callSortDirection =
      column === 'wait' ||
      column === 'priority' ||
      column === 'transfers' ||
      column === 'connected' ||
      column === 'connectedDuration'
        ? 'desc'
        : 'asc'
  }

  private toggleAgentSort(column: AgentSortColumn) {
    if (this.agentSortColumn === column) {
      this.agentSortDirection = this.agentSortDirection === 'asc' ? 'desc' : 'asc'
      return
    }

    this.agentSortColumn = column
    this.agentSortDirection = 'asc'
  }

  private toggleConnectedCallSort(column: CallSortColumn) {
    if (this.connectedCallSortColumn === column) {
      this.connectedCallSortDirection =
        this.connectedCallSortDirection === 'asc' ? 'desc' : 'asc'
      return
    }

    this.connectedCallSortColumn = column
    this.connectedCallSortDirection =
      column === 'wait' ||
      column === 'priority' ||
      column === 'transfers' ||
      column === 'connected' ||
      column === 'connectedDuration'
        ? 'desc'
        : 'asc'
  }

  private getSortIndicator(active: boolean, direction: 'asc' | 'desc') {
    if (!active) {
      return ''
    }

    return direction === 'asc' ? ' ▲' : ' ▼'
  }

  private compareValues(
    left: string | number | undefined,
    right: string | number | undefined,
    direction: 'asc' | 'desc'
  ) {
    const leftValue = left ?? ''
    const rightValue = right ?? ''

    let result = 0
    if (typeof leftValue === 'number' && typeof rightValue === 'number') {
      result = leftValue - rightValue
    } else {
      result = String(leftValue).localeCompare(String(rightValue), undefined, {
        numeric: true,
        sensitivity: 'base',
      })
    }

    return direction === 'asc' ? result : -result
  }

  private getSortedTasks() {
    return this.sortTasks(this.tasks, this.callSortColumn, this.callSortDirection)
  }

  private getSortedConnectedTasks() {
    return this.sortTasks(
      this.connectedTasks,
      this.connectedCallSortColumn,
      this.connectedCallSortDirection
    )
  }

  private sortTasks(
    tasks: QueueTask[],
    sortColumn: CallSortColumn,
    sortDirection: 'asc' | 'desc'
  ) {
    return [...tasks].sort((left, right) => {
      let leftValue: string | number | undefined
      let rightValue: string | number | undefined

      switch (sortColumn) {
        case 'wait':
          leftValue = left.waitTimeMs
          rightValue = right.waitTimeMs
          break
        case 'queue':
          leftValue = left.lastQueue?.name
          rightValue = right.lastQueue?.name
          break
        case 'ani':
          leftValue = left.origin || left.customer?.phoneNumber
          rightValue = right.origin || right.customer?.phoneNumber
          break
        case 'customer':
          leftValue = left.customer?.name || left.customer?.phoneNumber || left.origin
          rightValue = right.customer?.name || right.customer?.phoneNumber || right.origin
          break
        case 'status':
          leftValue = left.status
          rightValue = right.status
          break
        case 'priority':
          leftValue = left.contactPriority
          rightValue = right.contactPriority
          break
        case 'transfers':
          leftValue = left.transferCount
          rightValue = right.transferCount
          break
        case 'connected':
          leftValue = left.connectedCount
          rightValue = right.connectedCount
          break
        case 'connectedDuration':
          leftValue = left.connectedDuration
          rightValue = right.connectedDuration
          break
        case 'requiredSkills':
          leftValue = this.formatRequiredSkills(left.requiredSkills)
          rightValue = this.formatRequiredSkills(right.requiredSkills)
          break
        case 'callback':
          leftValue = left.callbackData?.callbackStatus
          rightValue = right.callbackData?.callbackStatus
          break
        case 'taskId':
          leftValue = left.id
          rightValue = right.id
          break
      }

      return this.compareValues(leftValue, rightValue, sortDirection)
    })
  }

  private getSortedAgents() {
    const filteredAgents = [...this.activeAgents].filter((agent) => {
      const idleCode = this.getAgentIdleCode(agent)
      const matchesState =
        this.agentStateFilter === 'all' ||
        this.getNormalizedAgentState(agent) === this.agentStateFilter
      const matchesTeam =
        this.agentTeamFilter === 'all' || String(agent.teamName ?? '') === this.agentTeamFilter
      const matchesIdleCode =
        this.agentIdleCodeFilter === 'all' || idleCode === this.agentIdleCodeFilter

      return matchesState && matchesTeam && matchesIdleCode
    })

    return filteredAgents.sort((left, right) => {
      let leftValue: string | number | undefined
      let rightValue: string | number | undefined

      switch (this.agentSortColumn) {
        case 'agentName':
          leftValue = left.agentName
          rightValue = right.agentName
          break
        case 'teamName':
          leftValue = left.teamName
          rightValue = right.teamName
          break
        case 'currentState':
          leftValue = this.getAgentStatePriority(left)
          rightValue = this.getAgentStatePriority(right)
          break
      }

      return this.compareValues(leftValue, rightValue, this.agentSortDirection)
    })
  }

  private renderCallSortButton(label: string, column: CallSortColumn) {
    return html`
      <button class="sort-button" @click=${() => this.toggleCallSort(column)}>
        ${label}${this.getSortIndicator(this.callSortColumn === column, this.callSortDirection)}
      </button>
    `
  }

  private renderAgentSortButton(label: string, column: AgentSortColumn) {
    return html`
      <button class="sort-button" @click=${() => this.toggleAgentSort(column)}>
        ${label}${this.getSortIndicator(
          this.agentSortColumn === column,
          this.agentSortDirection
        )}
      </button>
    `
  }

  private renderConnectedCallSortButton(label: string, column: CallSortColumn) {
    return html`
      <button class="sort-button" @click=${() => this.toggleConnectedCallSort(column)}>
        ${label}${this.getSortIndicator(
          this.connectedCallSortColumn === column,
          this.connectedCallSortDirection
        )}
      </button>
    `
  }

  private getTeamOptions() {
    return [...new Set(this.activeAgents.map((agent) => String(agent.teamName ?? '').trim()))]
      .filter((team) => team.length > 0)
      .sort((left, right) => left.localeCompare(right, undefined, { sensitivity: 'base' }))
  }

  private getAgentIdleCode(agent: ActiveAgentSession) {
    return String(this.getPrimaryChannel(agent)?.idleCodeName ?? '').trim()
  }

  private getIdleCodeOptions() {
    return [...new Set(this.activeAgents.map((agent) => this.getAgentIdleCode(agent)))]
      .filter((idleCode) => idleCode.length > 0)
      .sort((left, right) => left.localeCompare(right, undefined, { sensitivity: 'base' }))
  }

  private normalizeAgentSkillList(agentSkills: unknown): Array<Record<string, unknown>> {
    if (Array.isArray(agentSkills)) {
      return agentSkills.filter(
        (skill): skill is Record<string, unknown> =>
          Boolean(skill) && typeof skill === 'object' && !Array.isArray(skill)
      )
    }

    if (typeof agentSkills === 'string') {
      try {
        const parsed = JSON.parse(agentSkills) as unknown
        return this.normalizeAgentSkillList(parsed)
      } catch {
        return []
      }
    }

    return []
  }

  private formatAgentSkillValue(value: unknown): string {
    if (Array.isArray(value)) {
      return value.map((item) => this.formatAgentSkillValue(item)).filter(Boolean).join(', ')
    }

    if (value === undefined || value === null) {
      return ''
    }

    if (typeof value === 'object') {
      try {
        return JSON.stringify(value)
      } catch {
        return String(value)
      }
    }

    return String(value).trim()
  }

  private getAgentSkillDisplayValue(skill: Record<string, unknown>): string {
    const preferredKeys = [
      'skillValue',
      'value',
      'skillValues',
      'values',
      'intVal',
      'boolVal',
      'strVal',
      'stringVal',
      'enumVal',
      'enumValues',
      'numberVal',
      'numericVal',
    ]

    for (const key of preferredKeys) {
      if (key in skill) {
        const formatted = this.formatAgentSkillValue(skill[key])
        if (formatted) {
          return formatted
        }
      }
    }

    const ignoredKeys = new Set([
      'skillName',
      'name',
      '__typename',
      'id',
      'skillId',
      'description',
    ])

    for (const [key, rawValue] of Object.entries(skill)) {
      if (ignoredKeys.has(key)) {
        continue
      }

      const formatted = this.formatAgentSkillValue(rawValue)
      if (!formatted) {
        continue
      }

      return `${key}: ${formatted}`
    }

    return ''
  }

  private getAgentSkillDescriptions(agent: ActiveAgentSession) {
    const skills = this.normalizeAgentSkillList(agent.agentSkills)

    if (skills.length === 0) {
      return ['No skills returned']
    }

    const descriptions = skills
      .map((skill) => {
        const name = String(skill.skillName ?? skill.name ?? '').trim()
        const value = this.getAgentSkillDisplayValue(skill)

        if (!name) {
          return value
        }

        if (!value) {
          return name
        }

        return `${name}: ${value}`
      })
      .filter((description): description is string => Boolean(description))

    return descriptions.length > 0 ? descriptions : ['No skills returned']
  }

  private get themeClass() {
    return this.darkmode === 'true' ? 'theme-dark' : 'theme-light'
  }

  render() {
    const sortedTasks = this.getSortedTasks()
    const sortedConnectedTasks = this.getSortedConnectedTasks()
    const sortedAgents = this.getSortedAgents()
    const teamOptions = this.getTeamOptions()
    const idleCodeOptions = this.getIdleCodeOptions()

    return html`
      <section class="table-shell ${this.themeClass}">
        <div class="toolbar">
          <button class="refresh" @click=${this.loadData} ?disabled=${this.loading}>
            ${this.loading ? 'Refreshing...' : 'Refresh'}
          </button>
          <div class="meta">
            <span>${this.tasks.length} calls</span>
            <span>${this.connectedTasks.length} connected calls</span>
            <span>${this.activeAgents.length} active agents</span>
            <span>${this.apiRequestCount} API requests</span>
            <span>Status: ${this.statusFilter}</span>
            <span>Updated: ${this.lastUpdated || 'Not yet loaded'}</span>
          </div>
        </div>

        <div class="content-scroll">
          ${this.error
            ? html`<p class="status error">${this.error}</p>`
            : html`
              <div class="table-wrap">
                <div class="section-title">Calls Waiting In Queue</div>
                <table>
                  <thead>
                    <tr>
                      <th>${this.renderCallSortButton('Minutes', 'wait')}</th>
                      <th>${this.renderCallSortButton('Queue', 'queue')}</th>
                      <th>${this.renderCallSortButton('Required Skills', 'requiredSkills')}</th>
                      <th>${this.renderCallSortButton('ANI', 'ani')}</th>
                      <th>DNIS</th>
                      <th>${this.renderCallSortButton('Customer', 'customer')}</th>
                      <th>${this.renderCallSortButton('Status', 'status')}</th>
                      <th>${this.renderCallSortButton('Priority', 'priority')}</th>
                      <th>${this.renderCallSortButton('Transfers', 'transfers')}</th>
                      <th>${this.renderCallSortButton('Connected', 'connected')}</th>
                      <th>${this.renderCallSortButton('Connected Duration', 'connectedDuration')}</th>
                      <th>${this.renderCallSortButton('Session ID', 'taskId')}</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${sortedTasks.length > 0
                      ? sortedTasks.map(
                          (task) => html`
                            <tr>
                              <td class="strong">${this.formatWait(task.waitTimeMs)}</td>
                              <td>${task.lastQueue?.name || 'Unknown'}</td>
                              <td>${this.formatRequiredSkills(task.requiredSkills)}</td>
                              <td>${task.origin || task.customer?.phoneNumber || 'Unknown'}</td>
                              <td>${task.destination || 'Unknown'}</td>
                              <td>${task.customer?.name || 'Unknown caller'}</td>
                              <td>${task.status || 'Unknown'}</td>
                              <td>${task.contactPriority ?? 0}</td>
                              <td>${task.transferCount ?? 0}</td>
                              <td>${task.connectedCount ?? 0}</td>
                              <td>${this.formatConnectedDuration(task.connectedDuration)}</td>
                              <td>
                                <button class="copy-id" @click=${() => this.copyText(task.id)}>
                                  <span class="mono">${task.id}</span>
                                </button>
                              </td>
                            </tr>
                          `
                        )
                      : html`
                          <tr>
                            <td colspan="12">
                              No active ${this.statusFilter} calls found.
                            </td>
                          </tr>
                        `}
                  </tbody>
                </table>
              </div>

              <div class="table-wrap">
                <div class="section-title">Connected Calls</div>
                <table>
                  <thead>
                    <tr>
                      <th>${this.renderConnectedCallSortButton('Minutes', 'wait')}</th>
                      <th>${this.renderConnectedCallSortButton('Queue', 'queue')}</th>
                      <th>Agent</th>
                      <th>${this.renderConnectedCallSortButton('Required Skills', 'requiredSkills')}</th>
                      <th>${this.renderConnectedCallSortButton('ANI', 'ani')}</th>
                      <th>DNIS</th>
                      <th>${this.renderConnectedCallSortButton('Customer', 'customer')}</th>
                      <th>${this.renderConnectedCallSortButton('Priority', 'priority')}</th>
                      <th>${this.renderConnectedCallSortButton('Transfers', 'transfers')}</th>
                      <th>${this.renderConnectedCallSortButton('Connected', 'connected')}</th>
                      <th>${this.renderConnectedCallSortButton('Session ID', 'taskId')}</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${sortedConnectedTasks.length > 0
                      ? sortedConnectedTasks.map(
                          (task) => html`
                            <tr>
                              <td class="strong">${this.formatWait(task.waitTimeMs)}</td>
                              <td>${task.lastQueue?.name || 'Unknown'}</td>
                              <td>${task.owner?.name || 'Unknown'}</td>
                              <td>${this.formatRequiredSkills(task.requiredSkills)}</td>
                              <td>${task.origin || task.customer?.phoneNumber || 'Unknown'}</td>
                              <td>${task.destination || 'Unknown'}</td>
                              <td>${task.customer?.name || 'Unknown caller'}</td>
                              <td>${task.contactPriority ?? 0}</td>
                              <td>${task.transferCount ?? 0}</td>
                              <td>${task.connectedCount ?? 0}</td>
                              <td>
                                <button class="copy-id" @click=${() => this.copyText(task.id)}>
                                  <span class="mono">${task.id}</span>
                                </button>
                              </td>
                            </tr>
                          `
                        )
                      : html`
                          <tr>
                            <td colspan="11">No active connected calls found.</td>
                          </tr>
                        `}
                  </tbody>
                </table>
              </div>

              <div class="table-wrap">
                <div class="section-header">
                  <div class="section-title">Active Agents</div>
                  <div class="filters">
                    <label class="filter-label">
                      State
                      <select class="filter-select" @change=${this.setAgentStateFilter}>
                        <option value="all" ?selected=${this.agentStateFilter === 'all'}>
                          All
                        </option>
                        <option
                          value="available"
                          ?selected=${this.agentStateFilter === 'available'}
                        >
                          Available
                        </option>
                        <option value="idle" ?selected=${this.agentStateFilter === 'idle'}>
                          Idle
                        </option>
                        <option value="rona" ?selected=${this.agentStateFilter === 'rona'}>
                          RONA
                        </option>
                        <option value="wrapup" ?selected=${this.agentStateFilter === 'wrapup'}>
                          Wrapup
                        </option>
                      </select>
                    </label>

                    <label class="filter-label">
                      Team
                      <select class="filter-select" @change=${this.setAgentTeamFilter}>
                        <option value="all" ?selected=${this.agentTeamFilter === 'all'}>
                          All
                        </option>
                        ${teamOptions.map(
                          (team) => html`
                            <option value=${team} ?selected=${this.agentTeamFilter === team}>
                              ${team}
                            </option>
                          `
                        )}
                      </select>
                    </label>

                    <label class="filter-label">
                      Idle Code
                      <select class="filter-select" @change=${this.setAgentIdleCodeFilter}>
                        <option value="all" ?selected=${this.agentIdleCodeFilter === 'all'}>
                          All
                        </option>
                        ${idleCodeOptions.map(
                          (idleCode) => html`
                            <option
                              value=${idleCode}
                              ?selected=${this.agentIdleCodeFilter === idleCode}
                            >
                              ${idleCode}
                            </option>
                          `
                        )}
                      </select>
                    </label>
                  </div>
                </div>
                <table>
                  <thead>
                    <tr>
                      <th>${this.renderAgentSortButton('Agent', 'agentName')}</th>
                      <th>${this.renderAgentSortButton('Team', 'teamName')}</th>
                      <th>${this.renderAgentSortButton('Current State', 'currentState')}</th>
                      <th>Idle Code</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${sortedAgents.length > 0
                      ? sortedAgents.map(
                          (agent) => {
                            const skillDescriptions = this.getAgentSkillDescriptions(agent)
                            return html`
                              <tr>
                                <td class="agent-name-cell" title=${skillDescriptions.join('\n')}>
                                  <div class="agent-skill-hover">
                                    <span>${agent.agentName || 'Unknown'}</span>
                                    <div class="agent-skill-tooltip">
                                      <div class="agent-skill-tooltip-title">Skills</div>
                                      ${skillDescriptions.map(
                                        (description) => html`
                                          <div class="agent-skill-tooltip-line">
                                            ${description}
                                          </div>
                                        `
                                      )}
                                    </div>
                                  </div>
                                </td>
                                <td>${agent.teamName || 'Unknown'}</td>
                                <td>${this.getPrimaryChannel(agent)?.currentState || 'Unknown'}</td>
                                <td>${this.getAgentIdleCode(agent) || '—'}</td>
                              </tr>
                            `
                          }
                        )
                      : html`
                          <tr>
                            <td colspan="4">No active agents found.</td>
                          </tr>
                        `}
                  </tbody>
                </table>
              </div>
            `}
        </div>
      </section>
    `
  }

  static styles = css`
    :host {
      display: block;
      height: 100%;
    }

    * {
      box-sizing: border-box;
    }

    .table-shell {
      display: flex;
      flex-direction: column;
      height: 100%;
      min-height: 0;
      padding: 12px;
      font-family:
        'Segoe UI',
        -apple-system,
        BlinkMacSystemFont,
        sans-serif;
      color: #102a43;
      background: #f4f7fb;
    }

    .theme-dark {
      color: #f0f4f8;
      background: #08131f;
    }

    .toolbar {
      flex: 0 0 auto;
      display: flex;
      gap: 16px;
      align-items: center;
      justify-content: space-between;
      margin-bottom: 12px;
    }

    .content-scroll {
      flex: 1 1 auto;
      min-height: 0;
      overflow: auto;
    }

    .meta {
      display: flex;
      gap: 16px;
      flex-wrap: wrap;
      font-size: 0.9rem;
    }

    .refresh {
      border: 0;
      border-radius: 8px;
      padding: 8px 14px;
      font: inherit;
      font-weight: 700;
      color: white;
      background: #0f62fe;
      cursor: pointer;
    }

    .refresh[disabled] {
      opacity: 0.7;
      cursor: wait;
    }

    .status,
    .table-wrap {
      border: 1px solid rgba(16, 42, 67, 0.1);
      border-radius: 10px;
      background: white;
    }

    .status {
      padding: 14px 16px;
    }

    .error {
      color: #b42318;
    }

    .theme-dark .error {
      color: #fda29b;
    }

    .table-wrap {
      overflow: auto;
      margin-bottom: 12px;
    }

    .section-title {
      padding: 12px 12px 0;
      font-size: 0.85rem;
      font-weight: 700;
      letter-spacing: 0.06em;
      text-transform: uppercase;
    }

    .section-header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      padding-right: 12px;
    }

    .filter-label {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      font-size: 0.85rem;
      font-weight: 600;
      padding-top: 10px;
    }

    .filters {
      display: flex;
      gap: 12px;
      flex-wrap: wrap;
    }

    .filter-select {
      border: 1px solid #bcccdc;
      border-radius: 6px;
      background: white;
      color: inherit;
      font: inherit;
      padding: 4px 8px;
    }

    table {
      width: 100%;
      min-width: 1200px;
      border-collapse: collapse;
    }

    th,
    td {
      padding: 10px 12px;
      text-align: left;
      border-bottom: 1px solid #d9e2ec;
      vertical-align: top;
      font-size: 0.9rem;
    }

    th {
      position: sticky;
      top: 0;
      z-index: 1;
      background: #eaf1fb;
      font-size: 0.78rem;
      letter-spacing: 0.06em;
      text-transform: uppercase;
    }

    .sort-button {
      padding: 0;
      border: 0;
      background: transparent;
      color: inherit;
      font: inherit;
      font-weight: inherit;
      letter-spacing: inherit;
      text-transform: inherit;
      cursor: pointer;
    }

    .copy-id {
      padding: 0;
      border: 0;
      background: transparent;
      color: inherit;
      font: inherit;
      cursor: copy;
    }

    .copy-id:hover .mono {
      text-decoration: underline;
    }

    tr:hover td {
      background: #f8fbff;
    }

    .theme-dark .status,
    .theme-dark .table-wrap {
      background: #102a43;
      border-color: rgba(255, 255, 255, 0.12);
    }

    .theme-dark th {
      background: #16324f;
    }

    .theme-dark .filter-select {
      background: #16324f;
      border-color: rgba(255, 255, 255, 0.16);
      color: inherit;
    }

    .theme-dark th,
    .theme-dark td {
      border-bottom-color: rgba(255, 255, 255, 0.1);
    }

    .theme-dark tr:hover td {
      background: #173752;
    }

    .strong {
      font-weight: 700;
    }

    .agent-skill-hover {
      position: relative;
      display: inline-flex;
      align-items: center;
    }

    .agent-skill-tooltip {
      position: absolute;
      left: 0;
      top: calc(100% + 8px);
      z-index: 20;
      display: none;
      min-width: 280px;
      max-width: 420px;
      padding: 10px 12px;
      border: 1px solid #bcccdc;
      border-radius: 8px;
      background: #102a43;
      color: #f0f4f8;
      box-shadow: 0 12px 24px rgba(16, 42, 67, 0.18);
      white-space: normal;
      text-transform: none;
      letter-spacing: normal;
    }

    .agent-skill-hover:hover .agent-skill-tooltip {
      display: block;
    }

    .agent-skill-tooltip-title {
      font-size: 0.78rem;
      font-weight: 700;
      margin-bottom: 6px;
      text-transform: uppercase;
      letter-spacing: 0.06em;
      color: #d9e2ec;
    }

    .agent-skill-tooltip-line {
      font-size: 0.82rem;
      line-height: 1.4;
    }

    .mono {
      font-family:
        'SFMono-Regular',
        Consolas,
        'Liberation Mono',
        Menlo,
        monospace;
      font-size: 0.82rem;
    }

    @media (max-width: 720px) {
      .table-shell {
        min-height: auto;
        padding: 8px;
      }

      .toolbar {
        flex-direction: column;
        align-items: flex-start;
      }
    }
  `
}

declare global {
  interface HTMLElementTagNameMap {
    'queue-threshold-dashboard': MyElement
  }
}
