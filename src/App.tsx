import {
  AlertTriangle,
  Check,
  ChevronDown,
  ChevronUp,
  Clipboard,
  Database,
  Download,
  FileText,
  Inbox,
  KeyRound,
  ListFilter,
  MailOpen,
  RefreshCcw,
  Search,
  ShieldAlert,
  ShieldCheck,
  Shuffle,
  Trash2,
  Upload,
} from 'lucide-react';
import { ChangeEvent, useEffect, useMemo, useRef, useState } from 'react';
import {
  MailboxGroup,
  MailboxRecord,
  ParsedMailboxFile,
  exportAvailableMailboxes,
  formatMailboxCredential,
  getInventoryStats,
  getMailboxGroups,
  isUnsoldMailbox,
  markMailboxAvailable,
  markMailboxPrepared,
  markMailboxUsed,
  parseMailboxText,
  pickRandomAvailable,
} from './lib/mailboxLedger';
import './styles.css';

const SAMPLE_TEXT = [
  'AliceExample1001@outlook.com----pass-one----client-a----refresh-token-a',
  'BrendaExample2002@outlook.es----pass-two----client-b----refresh-token-b',
  'CarlosExample3003@outlook.com----pass-three----client-c----refresh-token-c 【6 月 19 日出 Perplexity】',
].join('\n');

const SHIPPING_NOTE = `Subject: Your Perplexity Pro Account - 12-Month Subscription

Hi there,

Thank you for your order! Your 12-month Perplexity Pro account is ready.

Please follow these 2 steps to get started. (Please read Step 1 carefully, as Microsoft will ask you to secure the email first!)

Step 1: Access your Outlook Inbox (Important!)
Go to website: outlook.live.com

Email: [Your Provided Email]
Password: [Your Provided Password]

What to expect during your first login:

Microsoft will immediately ask you to add a Recovery Email or Phone Number to secure the account.

Please enter your own personal email or phone number and complete Microsoft's verification.

Once completed, you will successfully enter the Outlook Inbox. Keep this inbox open!

Step 2: Log in to Perplexity
Go to website: perplexity.ai

Click "Sign in" -> "Continue with Email".

Enter the same email address: [Your Provided Email]

Perplexity will send a verification code/link to your Outlook inbox.

Go back to your Outlook inbox, copy the code (or click the link), and your 12-month Pro subscription is ready to use!

Tip: After logging into Outlook in Step 1, you can also change the password at account.microsoft.com to make it 100% yours.

If you love our service, a quick positive feedback on the order page would mean the world to us!

We also offer premium access to Gemini Pro, SuperGrok, Mobbin.com Pro... at unbeatable prices. Feel free to message us anytime if you need anything else!

Best regards,
Kiko`;

const CDK_NOTE = `Perplexity Topup Delivery
====================

Perplexity Redemption System --- Operation Guide (English)

Your 1-Year Pro Activation CDK:

[see order details]

QUICK 4-STEP SETUP (No Credit Card Required):

Step 1: Log in to Perplexity

In your browser, sign in to your target Perplexity account first.
Step 2: Get your User Info JSON

Open the official Perplexity session info page: perplexity.ai/api/auth/session (Copy this link into your browser)
Copy the entire text showing on that page (this is your User Info JSON).
Step 3: Fill in redemption information

Open our safe redemption portal: shopperplexity.top
In "Enter CDK Code", paste your CDK key provided above.
In "Paste User Info JSON", paste the entire text you copied in Step 2.
Step 4: Start redemption and confirm

Click "Start Redemption" to validate.
Check the account details shown, and click "Confirm" to complete the 1-Year Pro upgrade!
Important Notes and Warranty:

Subscription Overwrite: If your account already has an active Pro subscription, this redemption may overwrite it. Please use an account with no active Pro plan.

100% Free Assist: Too complicated for you? Don't worry! Just send us a message right here in the G2G Chat. We are online and always happy to complete the setup for you in 2 minutes!

Regards, Kiko`;

interface MailboxMessage {
  folder: string;
  uid: number | string;
  from: string;
  subject: string;
  date: string;
  body: string;
  otp: string;
}

interface DeviceCodeInfo {
  email: string;
  clientId: string;
  deviceCode: string;
  userCode: string;
  verificationUri: string;
  verificationUriComplete: string;
  expiresIn: number;
  interval: number;
}

type InventoryView = 'unsold' | 'available' | 'prepared';

function App() {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const detailPanelRef = useRef<HTMLDivElement>(null);
  const [records, setRecords] = useState<MailboxRecord[]>([]);
  const [parseErrors, setParseErrors] = useState<ReturnType<typeof parseMailboxText>['errors']>([]);
  const [selectedGroupKey, setSelectedGroupKey] = useState<string>('all');
  const [selectedMailboxId, setSelectedMailboxId] = useState<string>('');
  const [expandedUsedRecordId, setExpandedUsedRecordId] = useState<string>('');
  const [inventoryView, setInventoryView] = useState<InventoryView>('unsold');
  const [remark, setRemark] = useState('');
  const [preparationService, setPreparationService] = useState('Perplexity');
  const [preparationRemark, setPreparationRemark] = useState('');
  const [search, setSearch] = useState('');
  const [copyStatus, setCopyStatus] = useState<string>('');
  const [syncStatus, setSyncStatus] = useState<string>('正在读取共享台账...');
  const [isBusy, setIsBusy] = useState(false);
  const [mailBusy, setMailBusy] = useState(false);
  const [batchBusy, setBatchBusy] = useState(false);
  const [mailMessages, setMailMessages] = useState<MailboxMessage[]>([]);
  const [mailKeyword, setMailKeyword] = useState('');
  const [mailFolder, setMailFolder] = useState<'ALL' | 'INBOX' | 'Junk'>('ALL');
  const [mailMaxCount, setMailMaxCount] = useState(10);
  const [mailStatus, setMailStatus] = useState('');
  const [oauthEmail, setOauthEmail] = useState('');
  const [oauthPassword, setOauthPassword] = useState('');
  const [oauthClientId, setOauthClientId] = useState('9e5f94bc-e8a4-4e73-b8be-63364c29d753');
  const [oauthInfo, setOauthInfo] = useState<DeviceCodeInfo | null>(null);
  const [oauthStatus, setOauthStatus] = useState('');
  const [oauthBusy, setOauthBusy] = useState(false);

  useEffect(() => {
    void refreshRecords();
  }, []);

  const stats = useMemo(() => getInventoryStats(records), [records]);
  const tokenStats = useMemo(() => ({
    healthy: records.filter((record) => isUnsoldMailbox(record) && record.tokenStatus === 'healthy').length,
    error: records.filter((record) => isUnsoldMailbox(record) && record.tokenStatus === 'error').length,
  }), [records]);
  const groups = useMemo(() => getMailboxGroups(records), [records]);
  const selectedGroup = useMemo(
    () => groups.find((group) => group.key === selectedGroupKey),
    [groups, selectedGroupKey],
  );

  const visibleUnsoldRecords = useMemo(() => {
    return records
      .filter(isUnsoldMailbox)
      .filter((record) => inventoryView === 'unsold' || record.status === inventoryView)
      .filter((record) => {
        if (!selectedGroup) {
          return true;
        }

        return record.firstLetter === selectedGroup.firstLetter && record.domain === selectedGroup.domain;
      })
      .filter((record) => {
        const needle = search.trim().toLowerCase();
        if (!needle) {
          return true;
        }

        return (
          record.email.toLowerCase().includes(needle) ||
          record.domain.includes(needle) ||
          (record.preparedFor ?? '').toLowerCase().includes(needle) ||
          (record.preparationRemark ?? '').toLowerCase().includes(needle)
        );
      })
      .sort((a, b) => a.email.localeCompare(b.email));
  }, [inventoryView, records, search, selectedGroup]);

  const usedRecords = useMemo(
    () =>
      records
        .filter((record) => record.status === 'used')
        .filter((record) => {
          const needle = search.trim().toLowerCase();
          if (!needle) {
            return true;
          }

          return (
            record.email.toLowerCase().includes(needle) ||
            record.remark.toLowerCase().includes(needle) ||
            record.domain.includes(needle)
          );
        })
        .sort((a, b) => (b.usedAt ?? '').localeCompare(a.usedAt ?? '') || a.email.localeCompare(b.email)),
    [records, search],
  );

  const selectedMailbox = useMemo(
    () => records.find((record) => record.id === selectedMailboxId && isUnsoldMailbox(record)),
    [records, selectedMailboxId],
  );

  useEffect(() => {
    if (selectedMailboxId && !selectedMailbox) {
      setSelectedMailboxId('');
      setRemark('');
      setPreparationService('Perplexity');
      setPreparationRemark('');
    }
  }, [selectedMailbox, selectedMailboxId]);

  useEffect(() => {
    if (selectedMailbox && inventoryView !== 'unsold' && selectedMailbox.status !== inventoryView) {
      setSelectedMailboxId('');
    }
  }, [inventoryView, selectedMailbox]);

  useEffect(() => {
    setPreparationService(selectedMailbox?.preparedFor || 'Perplexity');
    setPreparationRemark(selectedMailbox?.preparationRemark || '');
  }, [selectedMailboxId, selectedMailbox?.preparedFor, selectedMailbox?.preparationRemark]);

  useEffect(() => {
    setMailMessages([]);
    setMailStatus('');
  }, [selectedMailboxId]);

  useEffect(() => {
    if (!selectedMailbox || !window.matchMedia('(max-width: 860px)').matches) {
      return;
    }

    window.setTimeout(() => {
      detailPanelRef.current?.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }, 0);
  }, [selectedMailbox]);

  async function refreshRecords() {
    setSyncStatus('正在读取共享台账...');

    try {
      const nextRecords = await fetchRecords();
      setRecords(nextRecords);
      setSyncStatus('共享台账已同步');
    } catch (error) {
      setSyncStatus(getErrorMessage(error));
    }
  }

  async function saveParsedRecords(parsed: ParsedMailboxFile) {
    setIsBusy(true);
    setSyncStatus('正在合并到共享台账...');

    try {
      const result = await importMailboxRecords(parsed.records);
      setRecords(result.records);
      setParseErrors(parsed.errors);
      setSelectedGroupKey('all');
      setSelectedMailboxId('');
      setRemark('');
      setSyncStatus(`导入完成：新增 ${result.created}，更新 ${result.updated}`);
    } catch (error) {
      setSyncStatus(getErrorMessage(error));
    } finally {
      setIsBusy(false);
    }
  }

  function importText(text: string) {
    const parsed = parseMailboxText(text);
    void saveParsedRecords(parsed);
  }

  function handleFileChange(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];

    if (!file) {
      return;
    }

    const reader = new FileReader();
    reader.onload = () => importText(String(reader.result ?? ''));
    reader.readAsText(file);
    event.target.value = '';
  }

  function handlePickRandom() {
    const picked = pickRandomAvailable(records, {
      firstLetter: selectedGroup?.firstLetter,
      domain: selectedGroup?.domain,
      status: inventoryView,
    });

    if (picked) {
      setSelectedMailboxId(picked.id);
      setRemark('');
    }
  }

  async function handleMarkPrepared() {
    if (!selectedMailbox || selectedMailbox.status !== 'available' || !preparationService.trim()) {
      return;
    }

    const optimisticRecords = markMailboxPrepared(
      records,
      selectedMailbox.id,
      preparationService,
      preparationRemark,
    );
    setRecords(optimisticRecords);
    setIsBusy(true);
    setSyncStatus('正在写入准备状态...');

    try {
      const nextRecords = await markRecordPrepared(
        selectedMailbox.id,
        preparationService,
        preparationRemark,
      );
      setRecords(nextRecords);
      setSyncStatus('已标记为准备状态');
    } catch (error) {
      await refreshRecords();
      setSyncStatus(getErrorMessage(error));
    } finally {
      setIsBusy(false);
    }
  }

  async function handleUndoPrepared() {
    if (!selectedMailbox || selectedMailbox.status !== 'prepared') {
      return;
    }
    if (!window.confirm('确定撤销准备状态吗？邮箱会回到“待准备”库存，准备记录将被清除。')) {
      return;
    }

    setRecords(markMailboxAvailable(records, selectedMailbox.id));
    setIsBusy(true);
    setSyncStatus('正在撤销准备状态...');

    try {
      const nextRecords = await unprepareRecord(selectedMailbox.id);
      setRecords(nextRecords);
      setSyncStatus('邮箱已回到待准备库存');
    } catch (error) {
      await refreshRecords();
      setSyncStatus(getErrorMessage(error));
    } finally {
      setIsBusy(false);
    }
  }

  async function handleMarkUsed() {
    if (!selectedMailbox || !remark.trim()) {
      return;
    }

    const optimisticRecords = markMailboxUsed(records, selectedMailbox.id, remark);
    setRecords(optimisticRecords);
    setIsBusy(true);
    setSyncStatus('正在写入共享台账...');

    try {
      const nextRecords = await markRecordUsed(selectedMailbox.id, remark);
      setRecords(nextRecords);
      setSelectedMailboxId('');
      setRemark('');
      setSyncStatus('共享台账已更新');
    } catch (error) {
      await refreshRecords();
      setSyncStatus(getErrorMessage(error));
    } finally {
      setIsBusy(false);
    }
  }

  async function handleFetchMail() {
    if (!selectedMailbox) return;
    setMailBusy(true);
    setMailStatus('正在连接 Outlook 并读取邮件...');
    try {
      const payload = await fetchRecordMail(selectedMailbox.id, {
        keyword: mailKeyword,
        folder: mailFolder,
        maxCount: mailMaxCount,
      });
      setRecords(payload.records);
      setMailMessages(payload.messages);
      setMailStatus(
        payload.messages.length
          ? `已读取 ${payload.messages.length} 封邮件`
          : mailKeyword.trim()
            ? '没有匹配邮件'
            : '当前所选文件夹暂无邮件',
      );
    } catch (error) {
      setMailStatus(getErrorMessage(error));
    } finally {
      setMailBusy(false);
    }
  }

  async function handleRefreshToken() {
    if (!selectedMailbox) return;
    setMailBusy(true);
    setMailStatus('正在刷新 Token...');
    try {
      const nextRecords = await refreshRecordToken(selectedMailbox.id);
      setRecords(nextRecords);
      setMailStatus('Token 刷新成功并已写回台账');
    } catch (error) {
      setMailStatus(getErrorMessage(error));
    } finally {
      setMailBusy(false);
    }
  }

  async function handleBatchRefresh() {
    if (!window.confirm('刷新 25 个从未检查或最久未检查的未售库存账号；已售账号会被自动排除。确定继续吗？')) return;
    setBatchBusy(true);
    setSyncStatus('正在分批刷新 25 个未售 Token...');
    try {
      const payload = await refreshAvailableTokens(25);
      setRecords(payload.records);
      setSyncStatus(`刷新完成：成功 ${payload.successCount}，失败 ${payload.failureCount}`);
    } catch (error) {
      setSyncStatus(getErrorMessage(error));
    } finally {
      setBatchBusy(false);
    }
  }

  async function handleStartOAuth() {
    setOauthBusy(true);
    setOauthStatus('正在生成 Microsoft 设备码...');
    try {
      const info = await startDeviceAuthorization(oauthEmail, oauthClientId);
      setOauthInfo(info);
      setOauthStatus('请在 Microsoft 页面完成登录和授权');
    } catch (error) {
      setOauthStatus(getErrorMessage(error));
    } finally {
      setOauthBusy(false);
    }
  }

  async function handleCheckOAuth() {
    if (!oauthInfo) return;
    setOauthBusy(true);
    setOauthStatus('正在检查授权状态...');
    try {
      const payload = await checkDeviceAuthorization(oauthInfo, oauthPassword);
      if (payload.pending) {
        setOauthStatus('Microsoft 尚未确认授权，请完成登录后再检查');
      } else {
        setRecords(payload.records ?? []);
        setOauthStatus('Token 已生成并保存到未售库存');
        setOauthInfo(null);
        setOauthEmail('');
        setOauthPassword('');
      }
    } catch (error) {
      setOauthStatus(getErrorMessage(error));
    } finally {
      setOauthBusy(false);
    }
  }

  function handleExport() {
    const content = exportAvailableMailboxes(records);
    const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'unsold-outlook-mailboxes.txt';
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 0);
  }

  async function handleCopyAvailable() {
    const copied = await writeClipboardText(exportAvailableMailboxes(records));
    setCopyStatus(copied ? '全部未售邮箱已复制' : '复制失败，请手动选中文本复制');
    window.setTimeout(() => setCopyStatus(''), 1600);
  }

  async function handleCopyShippingNote() {
    const copied = await writeClipboardText(SHIPPING_NOTE);
    setCopyStatus(copied ? '发货说明已复制' : '复制失败，请手动选中文本复制');
    window.setTimeout(() => setCopyStatus(''), 1600);
  }

  async function handleCopyCdkNote() {
    const copied = await writeClipboardText(CDK_NOTE);
    setCopyStatus(copied ? 'CDK 声明已复制' : '复制失败，请手动选中文本复制');
    window.setTimeout(() => setCopyStatus(''), 1600);
  }

  async function handleClear() {
    if (records.length !== 0 && !window.confirm('确定清空服务器共享台账吗？所有设备都会同步变为空。')) {
      return;
    }

    setIsBusy(true);
    setSyncStatus('正在清空共享台账...');
    try {
      const nextRecords = await replaceRecords([]);
      setRecords(nextRecords);
      setParseErrors([]);
      setSelectedMailboxId('');
      setSyncStatus('共享台账已清空');
    } catch (error) {
      setSyncStatus(getErrorMessage(error));
    } finally {
      setIsBusy(false);
    }
  }

  async function copyText(label: string, value: string) {
    const copied = await writeClipboardText(value);
    setCopyStatus(copied ? `${label}已复制` : '复制失败，请手动选中文本复制');
    window.setTimeout(() => setCopyStatus(''), 1600);
  }

  return (
    <main className="app-shell">
      <section className="topbar">
        <div>
          <h1>Outlook 邮箱运营台</h1>
          <p>库存准备、发货、OAuth Token 与完整收件箱，共用一份账号台账。</p>
        </div>
        <div className="topbar-actions">
          <input ref={fileInputRef} type="file" accept=".txt,text/plain" onChange={handleFileChange} hidden />
          <button className="button primary" onClick={() => fileInputRef.current?.click()}>
            <Upload size={17} />
            导入 TXT
          </button>
          <button className="button" onClick={handleExport} disabled={stats.unsold === 0 || isBusy}>
            <Download size={17} />
            导出未售邮箱
          </button>
          <button className="button" onClick={handleCopyAvailable} disabled={stats.unsold === 0 || isBusy}>
            <Clipboard size={17} />
            复制未售邮箱
          </button>
          <button className="button" onClick={handleCopyShippingNote}>
            <Clipboard size={17} />
            复制发货说明
          </button>
          <button className="button" onClick={handleCopyCdkNote}>
            <Clipboard size={17} />
            复制 CDK 声明
          </button>
          <button className="button" onClick={handleBatchRefresh} disabled={batchBusy || isBusy || stats.unsold === 0}>
            <RefreshCcw size={17} />
            刷新 25 个未售 Token
          </button>
          <button className="icon-button danger" onClick={handleClear} aria-label="清空共享台账" title="清空共享台账" disabled={isBusy}>
            <Trash2 size={18} />
          </button>
        </div>
      </section>

      <section className="stats-grid" aria-label="库存概览">
        <StatCard icon={<Database size={19} />} label="全部邮箱" value={stats.total} />
        <StatCard icon={<Inbox size={19} />} label="待准备" value={stats.available} />
        <StatCard icon={<ShieldCheck size={19} />} label="已准备" value={stats.prepared} tone="prepared" />
        <StatCard icon={<Check size={19} />} label="已使用" value={stats.used} tone="used" />
        <StatCard icon={<FileText size={19} />} label="outlook.com / .es" value={`${stats.outlookCom} / ${stats.outlookEs}`} />
        <StatCard icon={<ShieldCheck size={19} />} label="Token 正常" value={tokenStats.healthy} tone="good" />
        <StatCard icon={<ShieldAlert size={19} />} label="Token 异常" value={tokenStats.error} tone="danger" />
      </section>

      <section className="toolbar">
        <div className="search-box">
          <Search size={17} />
          <input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="搜索邮箱、后缀、准备服务或备注" />
        </div>
        <button className="button subtle" onClick={() => importText(SAMPLE_TEXT)} disabled={isBusy}>
          <RefreshCcw size={16} />
          载入示例
        </button>
        <button className="button subtle" onClick={() => void refreshRecords()} disabled={isBusy}>
          <RefreshCcw size={16} />
          同步
        </button>
        <span className="copy-status" role="status" aria-live="polite">{syncStatus}</span>
        {copyStatus ? <span className="copy-status" role="status" aria-live="polite">{copyStatus}</span> : null}
      </section>

      <details className="oauth-panel">
        <summary>为一个 Outlook 邮箱生成并保存 Token</summary>
        <div className="oauth-grid">
          <label>
            <span>邮箱</span>
            <input type="email" value={oauthEmail} onChange={(event) => setOauthEmail(event.target.value)} placeholder="user@outlook.com" />
          </label>
          <label>
            <span>初始密码（可选）</span>
            <input type="password" value={oauthPassword} onChange={(event) => setOauthPassword(event.target.value)} autoComplete="off" />
          </label>
          <label>
            <span>client_id</span>
            <input value={oauthClientId} onChange={(event) => setOauthClientId(event.target.value)} />
          </label>
          <button className="button" type="button" onClick={handleStartOAuth} disabled={oauthBusy || !oauthEmail.trim()}>
            生成设备码
          </button>
        </div>
        {oauthInfo ? (
          <div className="device-code-box">
            <a href={oauthInfo.verificationUriComplete || oauthInfo.verificationUri} target="_blank" rel="noreferrer">打开 Microsoft 授权页面</a>
            <code>{oauthInfo.userCode}</code>
            <button className="button primary" type="button" onClick={handleCheckOAuth} disabled={oauthBusy}>授权完成，检查并保存</button>
          </div>
        ) : null}
        {oauthStatus ? <p className="mail-status">{oauthStatus}</p> : null}
      </details>

      {records.length === 0 ? (
        <EmptyState onImport={() => fileInputRef.current?.click()} onSample={() => importText(SAMPLE_TEXT)} />
      ) : (
        <section className="workspace">
          <aside className="group-panel">
            <div className="panel-title">
              <ListFilter size={18} />
              <h2>分组</h2>
            </div>
            <GroupButton
              active={selectedGroupKey === 'all'}
              label="全部未售"
              detail={`待准备 ${stats.available} / 已准备 ${stats.prepared}`}
              onClick={() => setSelectedGroupKey('all')}
            />
            <div className="group-list">
              {groups.map((group) => (
                <GroupButton
                  key={group.key}
                  active={selectedGroupKey === group.key}
                  label={`${group.firstLetter} · ${group.domain}`}
                  detail={`待准备 ${group.availableCount} / 已准备 ${group.preparedCount} / 已用 ${group.usedCount}`}
                  disabled={group.availableCount + group.preparedCount === 0}
                  onClick={() => setSelectedGroupKey(group.key)}
                />
              ))}
            </div>
          </aside>

          <section className="available-panel">
            <div className="panel-header">
              <div>
                <h2>未售邮箱池</h2>
                <p>{selectedGroup ? `${selectedGroup.firstLetter} · ${selectedGroup.domain}` : '全部分组'}，当前显示 {visibleUnsoldRecords.length} 个</p>
              </div>
              <div className="panel-header-actions">
                <div className="inventory-tabs" aria-label="库存状态筛选">
                  <button className={inventoryView === 'unsold' ? 'active' : ''} type="button" onClick={() => setInventoryView('unsold')}>全部未售</button>
                  <button className={inventoryView === 'available' ? 'active' : ''} type="button" onClick={() => setInventoryView('available')}>待准备</button>
                  <button className={inventoryView === 'prepared' ? 'active' : ''} type="button" onClick={() => setInventoryView('prepared')}>已准备</button>
                </div>
                <button className="button primary" onClick={handlePickRandom} disabled={visibleUnsoldRecords.length === 0 || isBusy}>
                  <Shuffle size={17} />
                  随机抽取
                </button>
              </div>
            </div>

            <div className={`split-grid ${selectedMailbox ? 'has-selection' : ''}`}>
              <div className="mailbox-list" aria-label="未售邮箱列表">
                {visibleUnsoldRecords.length === 0 ? (
                  <div className="soft-empty">当前筛选下没有邮箱。</div>
                ) : (
                  visibleUnsoldRecords.slice(0, 80).map((record) => (
                    <button
                      key={record.id}
                      className={`mailbox-row ${selectedMailboxId === record.id ? 'selected' : ''}`}
                      onClick={() => setSelectedMailboxId(record.id)}
                    >
                      <span>{record.email}</span>
                      <small>{record.domain} · {record.status === 'prepared' ? `已准备：${record.preparedFor}` : '待准备'}</small>
                    </button>
                  ))
                )}
              </div>

              <div className="detail-panel" ref={detailPanelRef}>
                {selectedMailbox ? (
                  <>
                    <div className="detail-heading">
                      <h3>{selectedMailbox.email}</h3>
                      <span>
                        {selectedMailbox.firstLetter} · {selectedMailbox.domain} · {selectedMailbox.status === 'prepared' ? `已准备：${selectedMailbox.preparedFor}` : '待准备'} · Token {formatTokenStatus(selectedMailbox)}
                      </span>
                    </div>

                    <section className="mail-tools" aria-label="邮件收取">
                      <div className="mail-tools-title">
                        <div>
                          <strong>收件箱</strong>
                          <span>读取最新邮件，关键词留空即显示全部</span>
                        </div>
                        <span className="mail-tools-hint">收件箱与垃圾箱均可检索</span>
                      </div>
                      <div className="mail-tool-row">
                        <div className="mail-filter-row">
                          <select value={mailFolder} onChange={(event) => setMailFolder(event.target.value as 'ALL' | 'INBOX' | 'Junk')} aria-label="邮件文件夹">
                            <option value="ALL">收件箱 + 垃圾箱</option>
                            <option value="INBOX">仅收件箱</option>
                            <option value="Junk">仅垃圾箱</option>
                          </select>
                          <input
                            value={mailKeyword}
                            onChange={(event) => setMailKeyword(event.target.value)}
                            placeholder="搜索主题、发件人或正文"
                            aria-label="邮件关键词"
                          />
                          <select value={mailMaxCount} onChange={(event) => setMailMaxCount(Number(event.target.value))} aria-label="邮件数量">
                            <option value={10}>10 封</option>
                            <option value={20}>20 封</option>
                            <option value={30}>30 封</option>
                            <option value={50}>50 封</option>
                          </select>
                        </div>
                        <div className="mail-action-row">
                          <button className="button primary" type="button" onClick={handleFetchMail} disabled={mailBusy}>
                            <MailOpen size={17} />
                            读取邮件
                          </button>
                          <button className="button" type="button" onClick={handleRefreshToken} disabled={mailBusy}>
                            <RefreshCcw size={17} />
                            刷新 Token
                          </button>
                        </div>
                      </div>
                      {mailStatus || selectedMailbox.tokenError ? (
                        <div className="mail-feedback">
                          {mailStatus ? <p className="mail-status">{mailStatus}</p> : null}
                          {selectedMailbox.tokenError ? <p className="token-error">{selectedMailbox.tokenError}</p> : null}
                        </div>
                      ) : null}
                      {mailMessages.length ? <MailResults messages={mailMessages} onCopy={copyText} /> : null}
                    </section>

                    <section className={`preparation-card ${selectedMailbox.status}`} aria-label="免费账号准备状态">
                      {selectedMailbox.status === 'available' ? (
                        <>
                          <div className="preparation-heading">
                            <div>
                              <strong>免费账号准备</strong>
                              <span>在第三方网站完成免费账号注册后，在这里确认并留档。</span>
                            </div>
                            <span className="status-pill pending">待准备</span>
                          </div>
                          <div className="preparation-form">
                            <label>
                              <span>服务</span>
                              <input
                                value={preparationService}
                                onChange={(event) => setPreparationService(event.target.value)}
                                placeholder="例如：Perplexity"
                                autoComplete="off"
                              />
                            </label>
                            <label>
                              <span>准备备注（可选）</span>
                              <input
                                value={preparationRemark}
                                onChange={(event) => setPreparationRemark(event.target.value)}
                                placeholder="例如：免费账号已注册并验证"
                                autoComplete="off"
                              />
                            </label>
                          </div>
                          <button className="button prepared full" type="button" onClick={handleMarkPrepared} disabled={!preparationService.trim() || isBusy}>
                            <ShieldCheck size={17} />
                            确认已注册，标记为准备状态
                          </button>
                        </>
                      ) : (
                        <>
                          <div className="preparation-heading">
                            <div>
                              <strong>{selectedMailbox.preparedFor} 免费账号已准备</strong>
                              <span>
                                {selectedMailbox.preparedAt ? new Date(selectedMailbox.preparedAt).toLocaleString('zh-CN') : '准备时间未知'}
                                {selectedMailbox.preparationRemark ? ` · ${selectedMailbox.preparationRemark}` : ''}
                              </span>
                            </div>
                            <span className="status-pill prepared">已准备</span>
                          </div>
                          <button className="button subtle" type="button" onClick={handleUndoPrepared} disabled={isBusy}>
                            撤销准备状态
                          </button>
                        </>
                      )}
                    </section>

                    <div className="mobile-action-grid" aria-label="手机发货操作">
                      <button className="button primary" type="button" onClick={() => copyText('用户名', selectedMailbox.email)}>
                        <Clipboard size={17} />
                        复制用户名
                      </button>
                      <button className="button primary" type="button" onClick={() => copyText('密码', selectedMailbox.password)}>
                        <Clipboard size={17} />
                        复制密码
                      </button>
                      <button className="button" type="button" onClick={() => copyText('完整输出', formatMailboxCredential(selectedMailbox))}>
                        <Clipboard size={17} />
                        复制完整输出
                      </button>
                      <button className="button" type="button" onClick={handleCopyShippingNote}>
                        <Clipboard size={17} />
                        复制发货说明
                      </button>
                    </div>
                    {copyStatus ? <span className="detail-copy-status" role="status" aria-live="polite">{copyStatus}</span> : null}

                    <label className="remark-field">
                      <span>发货 / 使用备注</span>
                      <textarea
                        value={remark}
                        onChange={(event) => setRemark(event.target.value)}
                        placeholder="例如：6 月 19 日出 Perplexity / 交给 Alex"
                        rows={4}
                        autoComplete="off"
                      />
                    </label>
                    <button className="button primary full mark-used-button" type="button" onClick={handleMarkUsed} disabled={!remark.trim() || isBusy}>
                      <Check size={17} />
                      标记为已使用（已售出）
                    </button>

                    <CredentialField label="用户名" value={selectedMailbox.email} onCopy={copyText} />
                    <CredentialField label="密码" value={selectedMailbox.password} onCopy={copyText} />
                    <CredentialField label="client_id" value={selectedMailbox.clientId} onCopy={copyText} />
                    <CredentialField label="refresh_token" value={selectedMailbox.refreshToken} onCopy={copyText} multiline />
                    <CompleteCredentialField value={formatMailboxCredential(selectedMailbox)} onCopy={copyText} />
                  </>
                ) : (
                  <div className="detail-empty">
                    <Shuffle size={32} />
                    <h3>还没有选中邮箱</h3>
                    <p>点击左侧列表里的邮箱，或直接随机抽取一个。</p>
                  </div>
                )}
              </div>
            </div>
          </section>

          <aside className="used-panel">
            <div className="panel-header compact">
              <div>
                <h2>已使用台账</h2>
                <p>{usedRecords.length} 条记录</p>
              </div>
            </div>
            <div className="used-list">
              {usedRecords.length === 0 ? (
                <div className="soft-empty">还没有已使用记录。</div>
              ) : (
                usedRecords.map((record) => {
                  const isExpanded = expandedUsedRecordId === record.id;
                  return (
                    <article className={`used-row ${isExpanded ? 'expanded' : ''}`} key={record.id}>
                      <div className="used-row-summary">
                        <div className="used-row-copy">
                          <strong>{record.email}</strong>
                          <span>{record.remark || '未填写备注'}</span>
                          {record.preparedFor ? <span className="used-preparation">售出前已准备：{record.preparedFor}</span> : null}
                          <small>{record.usedAt ? new Date(record.usedAt).toLocaleString('zh-CN') : `原 TXT 第 ${record.sourceLineNumber} 行`}</small>
                        </div>
                        <button
                          className="button mini used-credential-toggle"
                          type="button"
                          aria-expanded={isExpanded}
                          aria-controls={`used-credentials-${record.id}`}
                          onClick={() => setExpandedUsedRecordId(isExpanded ? '' : record.id)}
                        >
                          <KeyRound size={15} />
                          {isExpanded ? '收起凭据' : '查看凭据'}
                          {isExpanded ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
                        </button>
                      </div>

                      {isExpanded ? (
                        <div className="used-credentials" id={`used-credentials-${record.id}`}>
                          <div className="used-credentials-notice">
                            <KeyRound size={17} />
                            <span>仅用于补发登录凭据；已售账号仍然禁止读取邮件和刷新 Token。</span>
                          </div>
                          {copyStatus ? <span className="used-copy-status" role="status" aria-live="polite">{copyStatus}</span> : null}
                          <div className="used-credential-grid">
                            <CredentialField label="用户名" value={record.email} onCopy={copyText} />
                            <CredentialField label="密码" value={record.password} onCopy={copyText} />
                            <CredentialField label="client_id" value={record.clientId} onCopy={copyText} />
                            <CredentialField label="refresh_token" value={record.refreshToken} onCopy={copyText} multiline />
                          </div>
                          <CompleteCredentialField value={formatMailboxCredential(record)} onCopy={copyText} />
                        </div>
                      ) : null}
                    </article>
                  );
                })
              )}
            </div>
          </aside>
        </section>
      )}

      {parseErrors.length > 0 ? (
        <section className="error-panel">
          <div className="panel-title">
            <AlertTriangle size={18} />
            <h2>导入时跳过的行</h2>
          </div>
          {parseErrors.map((error) => (
            <div className="error-row" key={`${error.lineNumber}-${error.line}`}>
              <strong>第 {error.lineNumber} 行</strong>
              <span>{error.reason}</span>
              <code>{error.line}</code>
            </div>
          ))}
        </section>
      ) : null}
    </main>
  );
}

interface StatCardProps {
  icon: React.ReactNode;
  label: string;
  value: number | string;
  tone?: 'good' | 'prepared' | 'used' | 'danger';
}

function StatCard({ icon, label, value, tone }: StatCardProps) {
  return (
    <article className={`stat-card ${tone ?? ''}`}>
      <div className="stat-icon">{icon}</div>
      <span>{label}</span>
      <strong>{value}</strong>
    </article>
  );
}

interface GroupButtonProps {
  active: boolean;
  label: string;
  detail: string;
  disabled?: boolean;
  onClick: () => void;
}

function GroupButton({ active, label, detail, disabled, onClick }: GroupButtonProps) {
  return (
    <button className={`group-button ${active ? 'active' : ''}`} disabled={disabled} onClick={onClick}>
      <span>{label}</span>
      <small>{detail}</small>
    </button>
  );
}

interface CredentialFieldProps {
  label: string;
  value: string;
  multiline?: boolean;
  onCopy: (label: string, value: string) => void;
}

function CredentialField({ label, value, multiline, onCopy }: CredentialFieldProps) {
  return (
    <div className={`credential-field ${multiline ? 'multiline' : ''}`}>
      <span>{label}</span>
      <code>{value}</code>
      <button className="icon-button" onClick={() => onCopy(label, value)} aria-label={`复制${label}`} title={`复制${label}`}>
        <Clipboard size={16} />
      </button>
    </div>
  );
}

function CompleteCredentialField({ value, onCopy }: { value: string; onCopy: (label: string, value: string) => void }) {
  return (
    <div className="complete-output">
      <div className="complete-output-header">
        <span>完整输出</span>
        <button className="button mini" onClick={() => onCopy('完整输出', value)}>
          <Clipboard size={15} />
          复制完整输出
        </button>
      </div>
      <code>{value}</code>
    </div>
  );
}

function MailResults({ messages, onCopy }: { messages: MailboxMessage[]; onCopy: (label: string, value: string) => void }) {
  return (
    <div className="mail-results">
      {messages.map((message) => (
        <article className="mail-card" key={`${message.folder}-${message.uid}`}>
          <div className="mail-card-head">
            <div>
              <strong>{message.subject || '(无主题)'}</strong>
              <span>{message.from || '(未知发件人)'}</span>
            </div>
            {message.otp ? (
              <button className="otp-button" type="button" onClick={() => onCopy('验证码', message.otp)}>
                {message.otp}
              </button>
            ) : null}
          </div>
          <div className="mail-card-meta">
            <span className="mail-folder-badge">{message.folder}</span>
            <time>{message.date ? new Date(message.date).toLocaleString('zh-CN') : '时间未知'}</time>
          </div>
          <details>
            <summary>查看邮件正文</summary>
            <pre>{message.body || '(无正文)'}</pre>
          </details>
        </article>
      ))}
    </div>
  );
}

function formatTokenStatus(record: MailboxRecord): string {
  if (record.tokenStatus === 'healthy') {
    return record.tokenCheckedAt ? `正常 · ${new Date(record.tokenCheckedAt).toLocaleDateString('zh-CN')}` : '正常';
  }
  if (record.tokenStatus === 'error') return '异常';
  return '未检测';
}

function EmptyState({ onImport, onSample }: { onImport: () => void; onSample: () => void }) {
  return (
    <section className="empty-state">
      <FileText size={42} />
      <h2>导入你的邮箱 TXT 开始管理</h2>
      <p>应用会识别行尾 `【备注】` 为已使用邮箱，其余进入待准备库存。数据保存在服务器共享台账，所有设备读取同一份记录。</p>
      <div className="empty-actions">
        <button className="button primary" onClick={onImport}>
          <Upload size={17} />
          选择 TXT 文件
        </button>
        <button className="button" onClick={onSample}>
          <RefreshCcw size={17} />
          载入示例数据
        </button>
      </div>
    </section>
  );
}

async function writeClipboardText(text: string): Promise<boolean> {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    return fallbackCopyText(text);
  }
}

function fallbackCopyText(text: string): boolean {
  const textarea = document.createElement('textarea');
  textarea.value = text;
  textarea.setAttribute('readonly', 'true');
  textarea.style.position = 'fixed';
  textarea.style.top = '0';
  textarea.style.left = '0';
  textarea.style.width = '1px';
  textarea.style.height = '1px';
  textarea.style.opacity = '0';
  textarea.style.fontSize = '16px';
  document.body.appendChild(textarea);
  textarea.focus();
  textarea.select();
  textarea.setSelectionRange(0, textarea.value.length);

  try {
    return document.execCommand('copy');
  } catch {
    return false;
  } finally {
    textarea.remove();
  }
}

async function fetchRecords(): Promise<MailboxRecord[]> {
  const response = await fetch('/api/records', { cache: 'no-store' });
  const payload = await readApiResponse(response);
  return payload.records;
}

async function replaceRecords(records: MailboxRecord[]): Promise<MailboxRecord[]> {
  const response = await fetch('/api/records', {
    method: 'PUT',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ records }),
  });
  const payload = await readApiResponse(response);
  return payload.records;
}

async function importMailboxRecords(records: MailboxRecord[]): Promise<{
  records: MailboxRecord[];
  created: number;
  updated: number;
}> {
  const response = await fetch('/api/records/import', {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ records }),
  });
  const payload = await response.json();
  if (!response.ok) throw new Error(payload.error || '共享台账导入失败');
  if (!Array.isArray(payload.records)) throw new Error('共享台账返回格式不正确');
  return payload;
}

async function markRecordUsed(id: string, nextRemark: string): Promise<MailboxRecord[]> {
  const response = await fetch(`/api/records/${encodeURIComponent(id)}/use`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ remark: nextRemark }),
  });
  const payload = await readApiResponse(response);
  return payload.records;
}

async function markRecordPrepared(
  id: string,
  preparedFor: string,
  preparationRemark: string,
): Promise<MailboxRecord[]> {
  const response = await fetch(`/api/records/${encodeURIComponent(id)}/prepare`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ preparedFor, preparationRemark }),
  });
  const payload = await readApiResponse(response);
  return payload.records;
}

async function unprepareRecord(id: string): Promise<MailboxRecord[]> {
  const response = await fetch(`/api/records/${encodeURIComponent(id)}/unprepare`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({}),
  });
  const payload = await readApiResponse(response);
  return payload.records;
}

async function fetchRecordMail(
  id: string,
  options: { keyword: string; folder: 'ALL' | 'INBOX' | 'Junk'; maxCount: number },
): Promise<{ records: MailboxRecord[]; messages: MailboxMessage[] }> {
  const response = await fetch(`/api/records/${encodeURIComponent(id)}/mail`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify(options),
  });
  const payload = await response.json();
  if (!response.ok) throw new Error(payload.error || '邮件读取失败');
  if (!Array.isArray(payload.records) || !Array.isArray(payload.messages)) throw new Error('邮件接口返回格式不正确');
  return payload;
}

async function refreshRecordToken(id: string): Promise<MailboxRecord[]> {
  const response = await fetch(`/api/records/${encodeURIComponent(id)}/refresh`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({}),
  });
  const payload = await readApiResponse(response);
  return payload.records;
}

async function refreshAvailableTokens(limit: number): Promise<{
  records: MailboxRecord[];
  successCount: number;
  failureCount: number;
}> {
  const response = await fetch('/api/records/refresh-available', {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ limit }),
  });
  const payload = await response.json();
  if (!response.ok) throw new Error(payload.error || '批量刷新失败');
  if (!Array.isArray(payload.records)) throw new Error('批量刷新接口返回格式不正确');
  return payload;
}

async function startDeviceAuthorization(email: string, clientId: string): Promise<DeviceCodeInfo> {
  const response = await fetch('/api/oauth/device-code', {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ email, clientId }),
  });
  const payload = await response.json();
  if (!response.ok) throw new Error(payload.error || '设备码生成失败');
  return payload;
}

async function checkDeviceAuthorization(info: DeviceCodeInfo, password: string): Promise<{
  pending: boolean;
  records?: MailboxRecord[];
}> {
  const response = await fetch('/api/oauth/poll', {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({
      email: info.email,
      password,
      clientId: info.clientId,
      deviceCode: info.deviceCode,
    }),
  });
  const payload = await response.json();
  if (!response.ok) throw new Error(payload.error || 'Microsoft 授权检查失败');
  return payload;
}

async function readApiResponse(response: Response): Promise<{ records: MailboxRecord[] }> {
  const payload = await response.json();

  if (!response.ok) {
    throw new Error(payload.error || '共享台账请求失败');
  }

  if (!Array.isArray(payload.records)) {
    throw new Error('共享台账返回格式不正确');
  }

  return payload;
}

function getErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }

  return '共享台账请求失败';
}

export default App;
