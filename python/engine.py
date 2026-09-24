"""Local qPCR analysis and plotting engine; JSON-lines stdin/stdout protocol."""
import base64
import io
import json
import math
import sys
import warnings

import numpy as np
import pandas as pd
from scipy import stats
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.ticker import ScalarFormatter
import seaborn as sns

sns.set_theme(style='whitegrid', context='paper', font='DejaVu Sans')
matplotlib.rcParams.update({'svg.fonttype': 'none', 'text.parse_math': False})


def finite(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


def analyze(p):
    base = p.get('base', 2)
    if not finite(base) or base <= 1:
        raise ValueError('Amplification base must be a finite number greater than 1.')
    genes, reference, calibrator = p.get('selectedGenes', []), p.get('refGene'), p.get('calibrator')
    if not genes or not reference or not calibrator:
        raise ValueError('Choose target genes, a reference gene, and a calibrator.')
    raw = [r for r in p.get('raw', []) if not r.get('omit') and finite(r.get('cq'))]
    if not raw:
        raise ValueError('No included numeric Ct values are available.')
    frame = pd.DataFrame(raw)
    mapping = p.get('sampleMap', {})
    frame['condition'] = frame['sample'].map(lambda s: mapping.get(s) or s)
    frame['gene_key'] = frame['target'].str.strip().str.casefold()
    grouped = frame.groupby(['trialId', 'condition', 'gene_key'], sort=False)['cq'].agg(['mean', 'count'])
    def get(trial, cond, gene):
        key = (trial, cond, gene.strip().casefold())
        return (float(grouped.loc[key, 'mean']), int(grouped.loc[key, 'count'])) if key in grouped.index else (None, 0)
    result = []
    for trial in p.get('trials', []):
        tid = trial['id']
        conditions = frame.loc[frame.trialId == tid, 'condition'].unique()
        for gene in genes:
            cal_target, _ = get(tid, calibrator, gene)
            cal_ref, _ = get(tid, calibrator, reference)
            cal_dct = cal_target - cal_ref if cal_target is not None and cal_ref is not None else None
            for cond in conditions:
                target, nt = get(tid, cond, gene)
                if not nt:
                    continue
                ref, nr = get(tid, cond, reference)
                dct = target - ref if ref is not None else None
                ddct = dct - cal_dct if dct is not None and cal_dct is not None else None
                log2fc = -ddct * math.log2(base) if ddct is not None else None
                with np.errstate(over='ignore', under='ignore'):
                    expression = float(np.power(base, -ddct)) if ddct is not None else None
                result.append(dict(trialId=tid, trial=trial['name'], gene=gene, condition=cond,
                                   targetMean=target, referenceMean=ref, dct=dct, calibratorDct=cal_dct,
                                   ddct=ddct, expression=expression if finite(expression) else None,
                                   log2fc=log2fc, nTarget=nt, nReference=nr))
    return {'results': result, 'engine': 'pandas / NumPy'}


def statistical_test(p):
    rows = [r for r in p.get('rows', []) if finite(r.get('value'))]
    conds = list(dict.fromkeys(r['condition'] for r in rows))
    kind = p.get('test', 'paired')
    if len(conds) < 2:
        raise ValueError('At least two conditions are required.')
    pair = p.get('pair') or conds[:2]
    groups = [[r['value'] for r in rows if r['condition'] == c] for c in conds]
    with warnings.catch_warnings():
        warnings.simplefilter('ignore', RuntimeWarning)
        if kind == 'anova':
            if any(len(g) < 2 for g in groups):
                raise ValueError('ANOVA requires at least two biological replicates per condition.')
            out = stats.f_oneway(*groups)
            name, detail = 'One-way ANOVA', f'F({len(groups)-1}, {sum(map(len,groups))-len(groups)}) · total n = {sum(map(len,groups))}'
        elif kind in ('paired', 'welch'):
            if len(pair) != 2 or any(c not in conds for c in pair) or pair[0] == pair[1]:
                raise ValueError('Choose two different available conditions.')
            maps = [{r['trialId']: r['value'] for r in rows if r['condition'] == c} for c in pair]
            if kind == 'paired':
                ids = [tid for tid in maps[0] if tid in maps[1]]
                if len(ids) < 2:
                    raise ValueError('Paired t-test requires at least 2 matched biological trials.')
                a, b = [[m[tid] for tid in ids] for m in maps]
                out = stats.ttest_rel(a, b)
                name, detail = 'Paired t-test', f'paired biological n = {len(ids)}'
            else:
                a, b = [list(m.values()) for m in maps]
                if min(len(a), len(b)) < 2:
                    raise ValueError("Welch's t-test requires at least 2 trials per condition.")
                out = stats.ttest_ind(a, b, equal_var=False)
                name, detail = "Welch's t-test", f'biological n = {len(a)} vs {len(b)}'
        else:
            raise ValueError('Unknown statistical test.')
    if not math.isfinite(float(out.pvalue)):
        raise ValueError('The test is undefined for these values (for example, identical constant groups).')
    return dict(name=name, detail=detail, p=float(out.pvalue),
                statistic=float(out.statistic) if math.isfinite(float(out.statistic)) else None,
                engine='SciPy', metric=p.get('metric'), overrides=bool(p.get('useOverrides')))


def render(p):
    settings = p.get('controls', {})
    metric, kind = settings.get('graphMetric', 'expression'), settings.get('graphType', 'bar')
    log = settings.get('yScale') == 'log'
    genes = p.get('genes', [])
    rows = [r for r in p.get('rows', []) if r['gene'] in genes and finite(r.get('value'))]
    omitted = sum(r['value'] <= 0 for r in rows) if log else 0
    if log:
        rows = [r for r in rows if r['value'] > 0]
    fmt = p.get('format', 'png')
    if fmt not in ('png', 'svg', 'pdf'):
        raise ValueError('Unsupported figure format.')
    colors = p.get('colors') or sns.color_palette('colorblind').as_hex()
    for c in colors:
        if not matplotlib.colors.is_color_like(c):
            raise ValueError('Invalid figure color.')
    opacity = min(1, max(.1, float(settings.get('barOpacity', .7))))
    point = min(8, max(2, float(settings.get('pointSize', 4.5))))
    width = min(820, max(320, float(p.get('width', 720)))) / 100
    fig, ax = plt.subplots(figsize=(width, 4.8), dpi=120)
    try:
        fig.subplots_adjust(left=.14, right=.97, top=.79, bottom=.29)
        ax.set_title(settings.get('graphTitle') or 'qPCR relative expression', fontsize=12, weight='bold', pad=35)
        labels = {'expression': 'Relative expression', 'log2fc': 'log₂ fold change', 'dct': 'ΔCt'}
        ax.set_ylabel(labels.get(metric, metric), fontsize=9)
        ax.yaxis.set_label_coords(.035, .53, transform=fig.transFigure)
        ax.set_axisbelow(True)
        if settings.get('showGrid', True):
            ax.grid(True, axis='y', color='#e6eaf0')
        else:
            ax.grid(False, axis='y')
        ax.grid(False, axis='x')
        sns.despine(ax=ax)
        if not rows:
            ax.text(.5, .5, 'No plottable values for this selection.', ha='center', va='center', transform=ax.transAxes)
            ax.set_xticks([])
        else:
            conds = list(dict.fromkeys(r['condition'] for r in rows))
            trials = list(dict.fromkeys(r['trialId'] for r in rows))
            names = {r['trialId']: r['trial'] for r in rows}
            categories = [(g,c) for g in genes for c in conds if any(r['gene']==g and r['condition']==c for r in rows)]
            grouped = kind == 'grouped'
            positions = {(g,c): gi + (ci-(len(conds)-1)/2)*.8/len(conds) for gi,g in enumerate(genes) for ci,c in enumerate(conds)} if grouped else {cat:i for i,cat in enumerate(categories)}
            bar_width = .72/len(conds) if grouped else .65
            for gi,(gene,cond) in enumerate(categories):
                subset = [r for r in rows if r['gene']==gene and r['condition']==cond]
                values = np.array([r['value'] for r in subset])
                x = positions[(gene,cond)]
                color = colors[(conds.index(cond) if grouped else genes.index(gene)) % len(colors)]
                avg = float(np.mean(values))
                err_type = settings.get('errorType', 'sd')
                err = float(np.std(values, ddof=1)) if len(values)>1 and err_type!='none' else 0
                if err_type=='sem':
                    err /= math.sqrt(len(values))
                if kind in ('bar','grouped'):
                    baseline = min(r['value'] for r in rows)*.8 if log else 0
                    ax.bar(x, avg-baseline, bottom=baseline, width=bar_width, color=color, edgecolor=color, alpha=opacity, linewidth=1)
                elif kind=='box':
                    ax.boxplot([values], positions=[x], widths=bar_width, patch_artist=True, showfliers=False,
                               manage_ticks=False, boxprops=dict(facecolor=color, alpha=opacity), medianprops=dict(color='#263a36'))
                elif kind=='points':
                    ax.hlines(avg, x-bar_width/3, x+bar_width/3, color=color, linewidth=1.5)
                if err and kind not in ('box','points'):
                    lower = min(err, avg*.999) if log else err
                    ax.errorbar(x, avg, yerr=[[lower],[err]], color='#45536a', capsize=3, fmt='none', linewidth=1)
                for r in subset:
                    ti = trials.index(r['trialId'])
                    offset = 0 if kind=='paired' or len(trials)==1 else (ti/(len(trials)-1)-.5)*bar_width*.6
                    ax.plot(x+offset, r['value'], 'o', markersize=point, markerfacecolor='white', markeredgecolor=colors[ti%len(colors)], markeredgewidth=1.3)
            if kind=='paired':
                for gene in genes:
                    for ti,tid in enumerate(trials):
                        points = [(positions[(gene,c)], next(r['value'] for r in rows if r['gene']==gene and r['condition']==c and r['trialId']==tid)) for c in conds if any(r['gene']==gene and r['condition']==c and r['trialId']==tid for r in rows)]
                        if len(points)>1:
                            ax.plot(*zip(*points), color=colors[ti%len(colors)], alpha=.65, linewidth=1, zorder=1)
            ticks = list(range(len(genes))) if grouped else list(range(len(categories)))
            ticklabels = genes if grouped else [g+' · '+c if len(genes)>1 else c for g,c in categories]
            # Bounded labels leave enough room for the 45-degree axis at fixed width.
            ticklabels = [s if len(s)<=32 else s[:31]+'…' for s in ticklabels]
            ax.set_xticks(ticks, ticklabels, rotation=45, ha='right', rotation_mode='anchor', fontsize=max(6,min(8,50/max(1,len(ticks)))))
            ax.set_xlim(-.6, max(ticks)+.6)
            if log:
                ax.set_yscale('log')
            elif metric=='expression':
                ax.set_ylim(bottom=0)
            threshold = settings.get('thresholdValue', 1)
            try:
                threshold = float(threshold)
            except (ValueError, TypeError):
                threshold = None
            if settings.get('showThreshold', True) and finite(threshold) and (not log or threshold>0):
                ax.axhline(threshold, color='#8290a5', linestyle=(0,(2,3)), linewidth=1)
                label = settings.get('thresholdLabel', '')
                if label:
                    ax.annotate(str(label)[:65], xy=(1,threshold), xycoords=('axes fraction','data'), xytext=(-3,4), textcoords='offset points', ha='right', fontsize=7, color='#657185')
            from matplotlib.lines import Line2D
            from matplotlib.patches import Patch
            handles = [Patch(facecolor=colors[i%len(colors)],alpha=opacity,label=c) for i,c in enumerate(conds)] if grouped else []
            handles += [Line2D([],[],marker='o',linestyle='none',markerfacecolor='white',markeredgecolor=colors[i%len(colors)],label=names[tid],markersize=4) for i,tid in enumerate(trials)]
            ax.legend(handles=handles,loc='lower left',bbox_to_anchor=(0,1.01),ncol=min(4,len(handles)),frameon=False,fontsize=6)
        ax.tick_params(axis='y', labelsize=8)
        data = io.BytesIO()
        fig.savefig(data, format=fmt, dpi=300 if p.get('export') else 120, facecolor='white', metadata={'Creator':'qPCR studio'} if fmt in ('pdf','svg') else None)
        return {'data':base64.b64encode(data.getvalue()).decode('ascii'), 'mime':{'png':'image/png','svg':'image/svg+xml','pdf':'application/pdf'}[fmt],
                'engine':'Matplotlib / Seaborn', 'omitted':omitted, 'width':round(width*100), 'labelAngle':45}
    finally:
        plt.close(fig)


def dispatch(action, payload):
    if action=='analyze': return analyze(payload)
    if action=='stats': return statistical_test(payload)
    if action=='render': return render(payload)
    if action=='health': return {'engine':'Python', 'numpy':np.__version__, 'pandas':pd.__version__, 'scipy':__import__('scipy').__version__, 'matplotlib':matplotlib.__version__, 'seaborn':sns.__version__}
    raise ValueError('Unknown action.')


if __name__=='__main__':
    for line in sys.stdin:
        request = {}
        try:
            request = json.loads(line)
            response = {'id':request['id'], 'result':dispatch(request['action'],request.get('payload',{}))}
        except Exception as exc:
            response = {'id':request.get('id'), 'error':str(exc)}
        print(json.dumps(response, allow_nan=False), flush=True)
