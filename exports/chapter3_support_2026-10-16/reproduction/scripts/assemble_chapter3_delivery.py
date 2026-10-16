"""Assemble reviewed chapter-3 artifacts; does not publish or fetch anything."""
from pathlib import Path
import csv
from decimal import Decimal as D
import hashlib
import json
import shutil
import zipfile

ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/'outputs/chapter3_support'
DEST=ROOT/'exports/chapter3_support_2026-10-16'


def read(path):
    with path.open(encoding='utf-8-sig',newline='') as f:return list(csv.DictReader(f))


def table(headers,rows):
    def clean(x):return str(x).replace('|','／').replace('\n',' ')
    return '\n'.join(['| '+' | '.join(map(clean,headers))+' |','|'+'---|'*len(headers)]+['| '+' | '.join(map(clean,r))+' |' for r in rows])+'\n'


def copy(src,dest):
    dest.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(src,dest)


def main():
    DEST.mkdir(parents=True,exist_ok=True)
    copy(OUT/'第三章完整写作支撑证据.xlsx',DEST/'08_chapter3_complete_evidence.xlsx')
    for p in OUT.glob('*.csv'):copy(p,DEST/'tables'/p.name)
    copy(OUT/'run_manifest.json',DEST/'run_manifest.json')
    figures=[(OUT/'全国风电装机规模.png','01_wind_capacity.png','全国风电装机规模'),(OUT/'电站项目转让价款调整.png','02_project_price_revision.png','项目转让价款调整'),(ROOT/'outputs/chapter3/利润与经营现金流.png','03_profit_cashflow.png','利润与经营现金流（各年原披露）'),(ROOT/'outputs/chapter3/2025营运资金现金流调节.png','04_working_capital_cashflow.png','2025营运资金现金流调节'),(ROOT/'outputs/chapter3_operating_cycle/存货结构与业务阶段.png','05_inventory_composition.png','存货结构与业务阶段')]
    for src,name,label in figures:copy(src,DEST/'figures'/name)
    for p in (OUT/'review').glob('*.png'):copy(p,DEST/'review'/p.name)
    for p in (ROOT/'exports/2026-10-16').glob('*.xlsx'):copy(p,DEST/'baseline'/p.name)
    for p in ['scripts/build_chapter3_support_pack.py','scripts/assemble_chapter3_delivery.py','tests/test_chapter3_support.py','requirements.txt']:
        copy(ROOT/p,DEST/'reproduction'/p)
    for p in ['data/processed/financial_statement_versions.csv','outputs/chapter3/annual_financial_signals.csv','outputs/chapter3/rd_inputs.csv']:
        copy(ROOT/p,DEST/'reproduction'/p)
    copy(ROOT/'outputs/chapter3/numeric_sources.csv',DEST/'tables/既有财务原值页码.csv')
    copy(ROOT/'outputs/chapter3_operating_cycle/contract_asset_full_scope.csv',DEST/'tables/完整合同资产.csv')
    copy(ROOT/'outputs/chapter3_operating_cycle/inventory_detail.csv',DEST/'tables/存货分项原值.csv')
    copy(ROOT/'outputs/chapter3_operating_cycle/disposal_cash_bridge.csv',DEST/'tables/处置现金汇总调节.csv')
    sources=read(OUT/'来源目录.csv')
    text=['# 来源及原文核验索引','','以下为26份来源材料，并非26项独立证据；公司回复与会计师回复有重叠。发行人PDF通过公开镜像取得，政府域名转载的新闻单独标注。访问及整理日为2026-10-16。','']
    for r in sources:
        url=r['url'];link=f'[{r["title"]}]({url})' if url.startswith('https://') else r['title']+'（用户上传版本）'
        text += [f'## {r["source_id"]}　{link}','',f'性质：{r["source_type"]}。网页／公告发布日：{r["publication_date"] or "未在本表推定"}；成文日：{r["document_date"] or "未另列"}。',f'原文件SHA256：`{r["sha256"]}`','']
    text += ['# 已核对原文与使用边界','']
    for r in read(OUT/'原文证据索引.csv'):
        text += [f'## {r["evidence_id"]}｜{r["source_id"]}｜{r["locator"]}','', '> '+r['quote'],'',r['interpretation']+'。边界：'+r['limit']+'。','']
    (DEST/'source_register.md').write_text('\n'.join(text)+'\n',encoding='utf-8')
    annual=read(OUT/'原披露财务序列.csv')
    money=['revenue','net_profit','cfo','inventory','receivables','short_loans','long_loans','capex']
    md=['# 第三章财务状况与指标附表','','金额为亿元；原始元值在工作簿与CSV中。各年当期披露尚未全期重述统一；2019仅作基期。长期借款是非流动部分，不等于全部银行借款。应收账款为净额，不含完整合同资产。','','## 表A：原披露金额序列','',table(['年度','营业收入','合并净利润','经营净现金流','存货','应收账款净额','短期借款','长期借款（非流动）','购建长期资产支出'],[[r['year']]+[f'{D(r[k])/D("1e8"):.2f}' for k in money] for r in annual]),'## 表B：描述性指标','']
    indicators=read(OUT/'描述性财务指标.csv');names=list(dict.fromkeys(r['indicator'] for r in indicators))
    indexed={(r['year'],r['indicator']):r for r in indicators}
    md += [table(['年度']+names,[[str(y)]+[f'{D(indexed[(str(y),n)]["value"]):.2f}' for n in names] for y in range(2019,2026)]),'资产负债率、销售净利率及经营现金流收入比单位为%；流动比率为倍；周转天数为天。周转按365天和同份年报期初期末净额均值计算，不代表合同逾期天数。','',table(['指标','公式'],[[n,next(r['formula'] for r in indicators if r['indicator']==n)] for n in names]),'## 表C：财务费用变动的会计项目桥接','','金额为亿元。减项按sign列处理，不能将费用增加全部归因于借款利息增加。','']
    fb=read(OUT/'财务费用变动桥接.csv')
    md += [table(['项目','2024','2025','对费用增加的调节额'],[[r['item'],f'{D(r["amount_2024_cny"])/D("1e8"):.4f}',f'{D(r["amount_2025_cny"])/D("1e8"):.4f}',f'{D(r["expense_change_cny"])/D("1e8"):.4f}'] for r in fb]),'来源：2025年报第202页，同表比较；合计费用由3.4244亿元增至5.1553亿元。','', '## 图表','']
    for _,name,label in figures:md += [f'### {label}','',f'![{label}](figures/{name})','']
    (DEST/'financial_tables.md').write_text('\n'.join(md)+'\n',encoding='utf-8')
    draft=(ROOT/'docs/chapter3_writing_support.md').read_text().replace('财务状况与指标表.md','[财务状况与指标表](financial_tables.md)')
    draft += '\n\n## 配套图表与原文\n\n[财务表及五幅图](financial_tables.md)｜[官方和公告原文索引](source_register.md)｜[争议与写作边界](evidence_limits.md)\n'
    (DEST/'09_chapter3_writing_support.md').write_text(draft,encoding='utf-8')
    conflicts=read(OUT/'争议与不可推断项.csv')
    (DEST/'evidence_limits.md').write_text('# 尚未解决的披露差异与写作边界\n\n这些差异不阻碍撰写已经得到支持的章节内容，但阻止下表所列更强结论。\n\n'+table(['编号','事项及来源','未解决问题','处理','限制'],[[r['id'],r['topic']+'；'+r['source'],r['issue'],r['treatment'],r['blocking']] for r in conflicts])+'\n## 信用期标示的PDF图像复核\n\n![汇总表，第15页](review/credit_summary.png)\n\n![应收细表，第27页](review/credit_detail.png)\n\n项目实际到账、合同到期明细及研发商业收益等尚未取得。不得以余额变动、交易价或预计完成日期代替。\n',encoding='utf-8')
    readiness=read(OUT/'章节准备度.csv')
    README='''# 第三章写作支撑包（2026-10-16）

当前材料足以按新版开题目录开展第三章写作：环境、企业概况、财务状况及四维风险信号均有对应证据。结论须限于可验证的事实和风险信号；未完成第四章赋权评分或第五章系统文本识别。

## 建议阅读顺序

1. [09：第三章完整写作支撑稿](09_chapter3_writing_support.md)：按3.1—3.4逐节组织，含分析段落、表格及结论边界。
2. [财务表及五幅图](financial_tables.md)：可在GitHub直接阅读，金额与指标公式明确。
3. [08：完整证据Excel](08_chapter3_complete_evidence.xlsx)：16张表；点击文件页面右上角下载图标获取。
4. [原文来源与摘录](source_register.md)：26份来源材料、46条核对摘录；五项政策、七年度统计，区分新闻转载、公司及会计师说明。
5. [披露差异及图像复核](evidence_limits.md)：五项争议，保留原文与适用边界。

## 关键发现

- 营运资金占用与现金实现是主线，但2025应收周转改善、存货周转基本持平，不能写作全面恶化。
- 两电站项目2024年转让价款合计下调约2.14亿元，公司说明涉及消纳、电价环境；减价不是已核实损失额。
- 洮南百强付款含具体条件，并单列1.11亿元资金偿还安排；尚未取得逐笔实际到账。
- 筹资分析区分借款期限、预计负债、拟申请授信和可用资金；财务费用已逐项桥接。
- 2026半年报只作未经审计的期后跟踪，未混入2020—2025主体样本。

## 计算与资料验证

25项自动测试通过；新增13项计算核验通过，46条摘录及26项金额／日期对保存文本完成核对。统计中的2023年新增并网量留空，未使用存量差额或媒体吊装量填补。逐项目实际收款匹配数仍为0，不把合同安排和金额吻合称为已到账。

## 原成果及复现

[baseline目录](baseline/)保留01—07工作簿作为前序快照；08和09是此次综合补充。旧快照中尚未核验官方政策等早期状态，以此次来源目录更新为准。

[tables目录](tables/)提供UTF-8 BOM CSV和页码索引；[run_manifest.json](run_manifest.json)记录文件哈希和检查范围；[reproduction目录](reproduction/)提供本轮代码、测试和指标计算输入。原始PDF和官网HTML未装入交付包，复现需按来源目录补齐相应工作区路径；用户上传的年报需使用相同SHA256版本，不能把同年不同版PDF无说明替换。图表字体使用Noto Sans CJK。

完整表格与段落仍需作者按学校体例统一表号、引用和文字；尚未证明的逾期、项目损失、政策因果效应及研发失败不应作为论文结论。

## 逐节准备度

'''+table(['章节','状态','证据','边界'],[[r['section'],r['status'],r['evidence'],r['limit']] for r in readiness])
    (DEST/'README.md').write_text(README,encoding='utf-8')
    checksum={str(p.relative_to(DEST)):hashlib.sha256(p.read_bytes()).hexdigest() for p in sorted(DEST.rglob('*')) if p.is_file() and p.name!='checksums.json'}
    (DEST/'checksums.json').write_text(json.dumps(checksum,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    archive=DEST.parent/'chapter3_support_2026-10-16.zip'
    with zipfile.ZipFile(archive,'w',zipfile.ZIP_DEFLATED) as z:
        for p in sorted(DEST.rglob('*')):
            if p.is_file():z.write(p,arcname=str(p.relative_to(DEST.parent)))
    with zipfile.ZipFile(archive) as z:assert z.testzip() is None
    print(json.dumps(dict(files=len(checksum)+1,zip_bytes=archive.stat().st_size,destination=str(DEST)),ensure_ascii=False))


if __name__=='__main__':main()
