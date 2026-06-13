// ========== 报告导出功能 ==========
// 使用 docx.js 生成真正的 Word 文档 (.docx)
// 使用 HTML + window.print() 生成 PDF

// 工具函数：格式化数字
function fmt(num, digits) {
    if (num === undefined || num === null || isNaN(num)) return '0';
    return Number(num).toFixed(digits);
}

// 工具函数：厘米转 twips（Word内部单位，1英寸=1440twips，2.54cm=1440twips）
function cmToTwip(cm) {
    return Math.round(cm / 2.54 * 1440);
}

// 读取界面上的某个输入值
function getInputValue(id) {
    const el = document.getElementById(id);
    if (!el) return '';
    return el.value;
}

// ========== 生成输入参数的完整列表（用于 Word 和 PDF） ==========
function getInputParamsAsList() {
    // 分 6 个分组展示所有输入参数
    return [
        {
            title: '1. 项目信息',
            items: [
                ['项目名称', getInputValue('project_name') || '（未填写）'],
                ['产品图号', getInputValue('drawing_number') || '（未填写）'],
                ['胶粘剂型号', getInputValue('adhesive_model') || '（未填写）'],
                ['窗户位置', getInputValue('window_position')]
            ]
        },
        {
            title: '2. 玻璃及安装参数',
            items: [
                ['玻璃质量', getInputValue('glass_mass') + ' kg'],
                ['玻璃面积', getInputValue('glass_area') + ' m²'],
                ['玻璃长度', getInputValue('glass_length') + ' mm'],
                ['玻璃宽度', getInputValue('glass_width') + ' mm'],
                ['安装角度', getInputValue('installation_angle') + ' °'],
                ['粘接宽度', getInputValue('bond_width') + ' mm'],
                ['胶层厚度', getInputValue('adhesive_thickness') + ' mm'],
                ['列车最大速度', getInputValue('train_speed') + ' km/h']
            ]
        },
        {
            title: '3. 环境与载荷参数',
            items: [
                ['风压', getInputValue('wind_pressure') + ' Pa'],
                ['纵向动载荷加速度', getInputValue('long_acceleration') + ' g'],
                ['横向动载荷加速度', getInputValue('trans_acceleration') + ' g'],
                ['垂直动载荷加速度', getInputValue('vert_acceleration') + ' g'],
                ['最低工作温度', getInputValue('min_temperature') + ' ℃'],
                ['最高工作温度', getInputValue('max_temperature') + ' ℃']
            ]
        },
        {
            title: '4. 材料性能参数',
            items: [
                ['玻璃热膨胀系数', getInputValue('glass_expansion') + ' /℃'],
                ['窗框热膨胀系数', getInputValue('frame_expansion') + ' /℃'],
                ['玻璃钢热膨胀系数', getInputValue('frp_expansion') + ' /℃'],
                ['初始抗拉强度', getInputValue('tensile_strength') + ' MPa'],
                ['初始剪切强度', getInputValue('shear_strength') + ' MPa'],
                ['允许最大剪切变形角', getInputValue('allow_shear_angle')]
            ]
        },
        {
            title: '5. 老化系数参数',
            items: [
                ['温度差异系数', getInputValue('temp_factor')],
                ['介质影响系数', getInputValue('medium_factor')],
                ['几何差异系数', getInputValue('geo_factor')],
                ['长期静态系数', getInputValue('long_static_factor')],
                ['动态载荷系数', getInputValue('dynamic_factor')],
                ['人工操作系数', getInputValue('human_factor')]
            ]
        },
        {
            title: '6. 安全系数要求',
            items: [
                ['剪切安全系数要求', '> ' + getInputValue('shear_safety_req')],
                ['拉伸安全系数要求', '> ' + getInputValue('tensile_safety_req')],
                ['户外高速工况', '> ' + getInputValue('front_outdoor_req')],
                ['隧道入口工况', '> ' + getInputValue('front_tunnel_entry_req')],
                ['隧道内稳定工况', '> ' + getInputValue('front_tunnel_stable_req')],
                ['隧道出口工况', '> ' + getInputValue('front_tunnel_exit_req')],
                ['会车叠加工况', '> ' + getInputValue('front_overtaking_req')]
            ]
        }
    ];
}

// ========== 导出 Word 报告（使用 docx.js） ==========
async function exportWordReport() {
    if (!window.calculationResults) {
        alert('请先点击"开始计算"！');
        return;
    }

    try {
        const inputs = window.calculationResults.inputs;
        const results = window.calculationResults.results;
        const steps = window.calculationResults.steps;

        const inputParams = getInputParamsAsList();
        const now = new Date();
        const dateStr = now.getFullYear() + ' 年 ' +
            String(now.getMonth() + 1).padStart(2, '0') + ' 月 ' +
            String(now.getDate()).padStart(2, '0') + ' 日 ' +
            String(now.getHours()).padStart(2, '0') + ':' +
            String(now.getMinutes()).padStart(2, '0') + ':' +
            String(now.getSeconds()).padStart(2, '0');

        const conditionNameMap = {
            'outdoor': '户外高速',
            'tunnel_entry': '隧道入口',
            'tunnel_stable': '隧道内稳定',
            'tunnel_exit': '隧道出口',
            'overtaking': '会车叠加'
        };

        const children = [];

        // 标题
        children.push(new docx.Paragraph({
            text: '胶粘剂性能评估报告',
            alignment: docx.AlignmentType.CENTER,
            spacing: { before: 0, after: 200 },
            font: '宋体',
            size: 28
        }));

        // ========== 1. 项目信息 ==========
        children.push(new docx.Paragraph({
            children: [new docx.TextRun({ text: inputParams[0].title, bold: true, size: 24 })],
            spacing: { before: 160, after: 60 },
            border: { bottom: { color: '333333', space: 4, style: docx.BorderStyle.SINGLE, size: 6 } }
        }));
        for (const [label, val] of inputParams[0].items) {
            children.push(new docx.Paragraph({
                children: [new docx.TextRun({ text: label + '  ' + val, size: 22 })],
                spacing: { after: 30 }
            }));
        }

        // ========== 2~6. 输入参数（分组表格） ==========
        children.push(new docx.Paragraph({
            children: [new docx.TextRun({ text: '输入参数', bold: true, size: 24 })],
            spacing: { before: 160, after: 60 },
            border: { bottom: { color: '333333', space: 4, style: docx.BorderStyle.SINGLE, size: 6 } }
        }));

        for (let g = 1; g < inputParams.length; g++) {
            const group = inputParams[g];

            children.push(new docx.Paragraph({
                children: [new docx.TextRun({ text: group.title, bold: true, size: 22 })],
                spacing: { before: 120, after: 40 }
            }));

            const tableRows = [];
            tableRows.push(new docx.TableRow({
                children: [
                    new docx.TableCell({
                        width: { size: 45, type: docx.WidthType.PERCENTAGE },
                        children: [new docx.Paragraph({
                            children: [new docx.TextRun({ text: '参数名称', bold: true, size: 20 })],
                            spacing: { before: 40, after: 40 }
                        })]
                    }),
                    new docx.TableCell({
                        width: { size: 55, type: docx.WidthType.PERCENTAGE },
                        children: [new docx.Paragraph({
                            children: [new docx.TextRun({ text: '数值', bold: true, size: 20 })],
                            spacing: { before: 40, after: 40 }
                        })]
                    })
                ]
            }));
            for (const [label, val] of group.items) {
                tableRows.push(new docx.TableRow({
                    children: [
                        new docx.TableCell({
                            width: { size: 45, type: docx.WidthType.PERCENTAGE },
                            children: [new docx.Paragraph({
                                children: [new docx.TextRun({ text: label, size: 20 })],
                                spacing: { before: 40, after: 40 }
                            })]
                        }),
                        new docx.TableCell({
                            width: { size: 55, type: docx.WidthType.PERCENTAGE },
                            children: [new docx.Paragraph({
                                children: [new docx.TextRun({ text: val, size: 20 })],
                                spacing: { before: 40, after: 40 }
                            })]
                        })
                    ]
                }));
            }
            children.push(new docx.Table({
                width: { size: 100, type: docx.WidthType.PERCENTAGE },
                rows: tableRows
            }));
        }

        // ========== 关键参数摘要 ==========
        children.push(new docx.Paragraph({
            children: [new docx.TextRun({ text: '关键参数摘要', bold: true, size: 24 })],
            spacing: { before: 160, after: 60 },
            border: { bottom: { color: '333333', space: 4, style: docx.BorderStyle.SINGLE, size: 6 } }
        }));
        children.push(new docx.Table({
            width: { size: 100, type: docx.WidthType.PERCENTAGE },
            rows: [
                new docx.TableRow({
                    children: [
                        new docx.TableCell({
                            width: { size: 45, type: docx.WidthType.PERCENTAGE },
                            children: [new docx.Paragraph({ children: [new docx.TextRun({ text: '参数名称', bold: true, size: 20 })], spacing: { before: 40, after: 40 } })]
                        }),
                        new docx.TableCell({
                            width: { size: 55, type: docx.WidthType.PERCENTAGE },
                            children: [new docx.Paragraph({ children: [new docx.TextRun({ text: '数值', bold: true, size: 20 })], spacing: { before: 40, after: 40 } })]
                        })
                    ]
                }),
                new docx.TableRow({ children: [
                    new docx.TableCell({ width: { size: 45, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ children: [new docx.TextRun({ text: '玻璃热膨胀系数', size: 20 })], spacing: { before: 40, after: 40 } })] }),
                    new docx.TableCell({ width: { size: 55, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ children: [new docx.TextRun({ text: inputs.glass_thermal_exp + ' /℃', size: 20 })], spacing: { before: 40, after: 40 } })] })
                ] }),
                new docx.TableRow({ children: [
                    new docx.TableCell({ width: { size: 45, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ children: [new docx.TextRun({ text: '窗框热膨胀系数', size: 20 })], spacing: { before: 40, after: 40 } })] }),
                    new docx.TableCell({ width: { size: 55, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ children: [new docx.TextRun({ text: inputs.frame_thermal_exp + ' /℃', size: 20 })], spacing: { before: 40, after: 40 } })] })
                ] }),
                new docx.TableRow({ children: [
                    new docx.TableCell({ width: { size: 45, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ children: [new docx.TextRun({ text: '玻璃钢热膨胀系数', size: 20 })], spacing: { before: 40, after: 40 } })] }),
                    new docx.TableCell({ width: { size: 55, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ children: [new docx.TextRun({ text: inputs.frp_thermal_exp + ' /℃', size: 20 })], spacing: { before: 40, after: 40 } })] })
                ] }),
                new docx.TableRow({ children: [
                    new docx.TableCell({ width: { size: 45, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ children: [new docx.TextRun({ text: '胶层面积', size: 20 })], spacing: { before: 40, after: 40 } })] }),
                    new docx.TableCell({ width: { size: 55, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ children: [new docx.TextRun({ text: fmt(results.bond_area, 2) + ' mm²', size: 20 })], spacing: { before: 40, after: 40 } })] })
                ] }),
                new docx.TableRow({ children: [
                    new docx.TableCell({ width: { size: 45, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ children: [new docx.TextRun({ text: '剪切应力', size: 20 })], spacing: { before: 40, after: 40 } })] }),
                    new docx.TableCell({ width: { size: 55, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ children: [new docx.TextRun({ text: fmt(results.tau_shear, 6) + ' MPa', size: 20 })], spacing: { before: 40, after: 40 } })] })
                ] }),
                new docx.TableRow({ children: [
                    new docx.TableCell({ width: { size: 45, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ children: [new docx.TextRun({ text: '拉伸应力', size: 20 })], spacing: { before: 40, after: 40 } })] }),
                    new docx.TableCell({ width: { size: 55, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ children: [new docx.TextRun({ text: fmt(results.sigma_tensile, 6) + ' MPa', size: 20 })], spacing: { before: 40, after: 40 } })] })
                ] }),
                new docx.TableRow({ children: [
                    new docx.TableCell({ width: { size: 45, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ children: [new docx.TextRun({ text: '总老化系数', size: 20 })], spacing: { before: 40, after: 40 } })] }),
                    new docx.TableCell({ width: { size: 55, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ children: [new docx.TextRun({ text: fmt(results.total_aging_factor, 6), size: 20 })], spacing: { before: 40, after: 40 } })] })
                ] })
            ]
        }));

        // ========== 评估结果 ==========
        children.push(new docx.Paragraph({
            children: [new docx.TextRun({ text: '评估结果', bold: true, size: 24 })],
            spacing: { before: 160, after: 60 },
            border: { bottom: { color: '333333', space: 4, style: docx.BorderStyle.SINGLE, size: 6 } }
        }));
        children.push(new docx.Paragraph({
            children: [new docx.TextRun({ text: '总体评估: ' + (results.is_safe ? '胶粘剂满足要求' : '胶粘剂不满足要求'), size: 22 })],
            spacing: { after: 40 }
        }));
        children.push(new docx.Paragraph({
            children: [new docx.TextRun({ text: '剪切安全系数: ' + fmt(results.shear_safety_factor, 2) + ' (要求: > ' + inputs.shear_req + ')', size: 22 })],
            spacing: { after: 40 }
        }));
        children.push(new docx.Paragraph({
            children: [new docx.TextRun({ text: '拉伸安全系数: ' + fmt(results.tensile_safety_factor, 2) + ' (要求: > ' + inputs.tensile_req + ')', size: 22 })],
            spacing: { after: 40 }
        }));
        children.push(new docx.Paragraph({
            children: [new docx.TextRun({ text: '剪切变形角: ' + fmt(results.tan_theta, 6) + ' (要求: ≤ ' + inputs.allow_shear_angle + ')', size: 22 })],
            spacing: { after: 40 }
        }));

        // ========== 隧道及会车工况 ==========
        children.push(new docx.Paragraph({
            children: [new docx.TextRun({ text: '隧道及会车工况安全系数', bold: true, size: 24 })],
            spacing: { before: 160, after: 60 },
            border: { bottom: { color: '333333', space: 4, style: docx.BorderStyle.SINGLE, size: 6 } }
        }));
        for (const [condition, factor] of Object.entries(results.condition_safety_factors)) {
            const req = results.condition_requirements[condition];
            const name = conditionNameMap[condition] || condition;
            const passed = results.condition_safety_status[condition];
            children.push(new docx.Paragraph({
                children: [new docx.TextRun({ text: name + ': ' + fmt(factor, 2) + ' (要求: > ' + req + ') ' + (passed ? '(满足)' : '(不满足)'), size: 22 })],
                spacing: { after: 40 }
            }));
        }

        // ========== 详细计算过程 ==========
        children.push(new docx.Paragraph({
            children: [new docx.TextRun({ text: '详细计算过程', bold: true, size: 24 })],
            spacing: { before: 200, after: 60 },
            border: { bottom: { color: '333333', space: 4, style: docx.BorderStyle.SINGLE, size: 6 } }
        }));
        for (const step of steps) {
            children.push(new docx.Paragraph({
                children: [new docx.TextRun({ text: step.title, bold: true, size: 22 })],
                spacing: { before: 120, after: 60 }
            }));
            for (const calc of step.calculations) {
                children.push(new docx.Paragraph({
                    children: [new docx.TextRun({ text: '公式: ' + calc.formula, size: 20, bold: true })],
                    spacing: { after: 30 }
                }));
                children.push(new docx.Paragraph({
                    children: [new docx.TextRun({ text: '计算: ' + calc.value, size: 20 })],
                    spacing: { after: 60 }
                }));
            }
        }

        // ========== 结论与建议 ==========
        children.push(new docx.Paragraph({
            children: [new docx.TextRun({ text: '结论与建议', bold: true, size: 24 })],
            spacing: { before: 200, after: 60 },
            border: { bottom: { color: '333333', space: 4, style: docx.BorderStyle.SINGLE, size: 6 } }
        }));
        children.push(new docx.Paragraph({
            children: [new docx.TextRun({
                text: (results.is_safe
                    ? '根据计算结果，该胶粘剂的各项性能指标均满足设计要求，包括隧道和会车工况，可以在本项目中使用。'
                    : '根据计算结果，该胶粘剂部分性能指标不满足设计要求，建议调整胶粘剂参数或更换更适合的胶粘剂类型。'),
                size: 22
            })],
            spacing: { after: 40 }
        }));

        // ========== 编审批 ==========
        children.push(new docx.Paragraph({
            children: [new docx.TextRun({ text: '编审批', bold: true, size: 24 })],
            spacing: { before: 200, after: 60 },
            border: { bottom: { color: '333333', space: 4, style: docx.BorderStyle.SINGLE, size: 6 } }
        }));
        children.push(new docx.Table({
            width: { size: 100, type: docx.WidthType.PERCENTAGE },
            rows: [
                new docx.TableRow({
                    children: [
                        new docx.TableCell({ width: { size: 30, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ children: [new docx.TextRun({ text: '', size: 20 })], spacing: { before: 60, after: 60 } })] }),
                        new docx.TableCell({ width: { size: 35, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ children: [new docx.TextRun({ text: '姓名', bold: true, size: 20 })], spacing: { before: 60, after: 60 } })] }),
                        new docx.TableCell({ width: { size: 35, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ children: [new docx.TextRun({ text: '日期', bold: true, size: 20 })], spacing: { before: 60, after: 60 } })] })
                    ]
                }),
                new docx.TableRow({
                    children: [
                        new docx.TableCell({ width: { size: 30, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ children: [new docx.TextRun({ text: '编制', size: 20 })], spacing: { before: 80, after: 80 } })] }),
                        new docx.TableCell({ width: { size: 35, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ children: [new docx.TextRun({ text: '', size: 20 })], spacing: { before: 80, after: 80 } })] }),
                        new docx.TableCell({ width: { size: 35, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ children: [new docx.TextRun({ text: '', size: 20 })], spacing: { before: 80, after: 80 } })] })
                    ]
                }),
                new docx.TableRow({
                    children: [
                        new docx.TableCell({ width: { size: 30, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ children: [new docx.TextRun({ text: '审核', size: 20 })], spacing: { before: 80, after: 80 } })] }),
                        new docx.TableCell({ width: { size: 35, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ children: [new docx.TextRun({ text: '', size: 20 })], spacing: { before: 80, after: 80 } })] }),
                        new docx.TableCell({ width: { size: 35, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ children: [new docx.TextRun({ text: '', size: 20 })], spacing: { before: 80, after: 80 } })] })
                    ]
                }),
                new docx.TableRow({
                    children: [
                        new docx.TableCell({ width: { size: 30, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ children: [new docx.TextRun({ text: '批准', size: 20 })], spacing: { before: 80, after: 80 } })] }),
                        new docx.TableCell({ width: { size: 35, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ children: [new docx.TextRun({ text: '', size: 20 })], spacing: { before: 80, after: 80 } })] }),
                        new docx.TableCell({ width: { size: 35, type: docx.WidthType.PERCENTAGE }, children: [new docx.Paragraph({ children: [new docx.TextRun({ text: '', size: 20 })], spacing: { before: 80, after: 80 } })] })
                    ]
                })
            ]
        }));

        children.push(new docx.Paragraph({
            children: [new docx.TextRun({ text: '报告生成日期: ' + dateStr, size: 20 })],
            spacing: { before: 200 },
            alignment: docx.AlignmentType.RIGHT
        }));

        // 创建文档（使用手动 cm 转 twip，2.54cm = 1 inch）
        const doc = new docx.Document({
            creator: '胶粘剂性能评估系统',
            title: '',
            features: { updateFields: true },
            sections: [{
                properties: {
                    page: {
                        size: {
                            width: cmToTwip(21),
                            height: cmToTwip(29.7)
                        },
                        margin: {
                            top: cmToTwip(2.54),
                            bottom: cmToTwip(2.54),
                            left: cmToTwip(2.54),
                            right: cmToTwip(2.54),
                            header: cmToTwip(0.5),
                            footer: cmToTwip(0.5)
                        }
                    }
                },
                headers: {
                    default: new docx.Header({ children: [] })
                },
                footers: {
                    default: new docx.Footer({ children: [] })
                },
                children: children
            }]
        });

        // 生成并下载
        const blob = await docx.Packer.toBlob(doc);
        const filename = '胶粘剂性能评估报告_' +
            now.getFullYear() +
            String(now.getMonth() + 1).padStart(2, '0') +
            String(now.getDate()).padStart(2, '0') +
            '.docx';
        saveAs(blob, filename);

    } catch (e) {
        console.error('导出Word失败:', e);
        alert('导出Word失败: ' + (e.message || e));
    }
}

// ========== 导出 PDF 报告（严格模仿 Word 排版：紧凑表格、分组标题、下划线分隔） ==========
async function exportPdfReport() {
    if (!window.calculationResults) {
        alert('请先点击"开始计算"！');
        return;
    }

    try {
        const inputs = window.calculationResults.inputs;
        const results = window.calculationResults.results;
        const steps = window.calculationResults.steps;

        const inputParams = getInputParamsAsList();
        const now = new Date();
        const dateStr = now.getFullYear() + ' 年 ' +
            String(now.getMonth() + 1).padStart(2, '0') + ' 月 ' +
            String(now.getDate()).padStart(2, '0') + ' 日 ' +
            String(now.getHours()).padStart(2, '0') + ':' +
            String(now.getMinutes()).padStart(2, '0') + ':' +
            String(now.getSeconds()).padStart(2, '0');

        const conditionNameMap = {
            'outdoor': '户外高速',
            'tunnel_entry': '隧道入口',
            'tunnel_stable': '隧道内稳定',
            'tunnel_exit': '隧道出口',
            'overtaking': '会车叠加'
        };

        // 构建参数表格 HTML
        let paramTablesHTML = '';
        for (let g = 1; g < inputParams.length; g++) {
            const group = inputParams[g];
            paramTablesHTML += '<div class="group-title">' + group.title + '</div>\n';
            paramTablesHTML += '<table>\n';
            paramTablesHTML += '<tr><th class="pname">参数名称</th><th class="pval">数值</th></tr>\n';
            for (const [label, val] of group.items) {
                paramTablesHTML += '<tr><td class="pname">' + label + '</td><td class="pval">' + val + '</td></tr>\n';
            }
            paramTablesHTML += '</table>\n';
        }

        // 隧道工况评估
        let conditionHTML = '';
        for (const [condition, factor] of Object.entries(results.condition_safety_factors)) {
            const req = results.condition_requirements[condition];
            const name = conditionNameMap[condition] || condition;
            const passed = results.condition_safety_status[condition];
            const mark = passed ? '(满足)' : '(不满足)';
            conditionHTML += '<div class="result-line">' + name + ': ' + fmt(factor, 2) + ' (要求: > ' + req + ') ' + mark + '</div>\n';
        }

        // 详细计算过程
        let stepsHTML = '';
        for (const step of steps) {
            stepsHTML += '<div class="step-title">' + step.title + '</div>\n';
            for (const calc of step.calculations) {
                stepsHTML += '<div class="calc-item formula">公式: ' + calc.formula + '</div>\n';
                stepsHTML += '<div class="calc-item value">计算: ' + calc.value + '</div>\n';
            }
        }

        // ========== 构建完整 HTML（严格模仿 Word 排版风格） ==========
        let html = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title> </title>
    <style>
        @page {
            size: A4;
            margin: 2.54cm;
        }
        @page :left {
            @top-left-corner { content: ""; }
            @top-left { content: ""; }
            @top-center { content: ""; }
            @top-right { content: ""; }
            @top-right-corner { content: ""; }
            @bottom-left-corner { content: ""; }
            @bottom-left { content: ""; }
            @bottom-center { content: ""; }
            @bottom-right { content: ""; }
            @bottom-right-corner { content: ""; }
        }
        @page :right {
            @top-left-corner { content: ""; }
            @top-left { content: ""; }
            @top-center { content: ""; }
            @top-right { content: ""; }
            @top-right-corner { content: ""; }
            @bottom-left-corner { content: ""; }
            @bottom-left { content: ""; }
            @bottom-center { content: ""; }
            @bottom-right { content: ""; }
            @bottom-right-corner { content: ""; }
        }
        body {
            font-family: "SimSun", "宋体", serif;
            line-height: 1.4;
            color: #000;
            font-size: 11pt;
            margin: 0;
        }
        .report-title {
            text-align: center;
            font-size: 16pt;
            font-weight: bold;
            margin: 0 0 12pt 0;
        }
        .section-title {
            font-size: 12pt;
            font-weight: bold;
            margin-top: 10pt;
            margin-bottom: 4pt;
            padding-bottom: 2pt;
            border-bottom: 1pt solid #333;
        }
        .group-title {
            font-size: 11pt;
            font-weight: bold;
            margin-top: 6pt;
            margin-bottom: 2pt;
        }
        .info-item {
            margin: 2pt 0;
            font-size: 10.5pt;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            margin: 3pt 0 4pt 0;
        }
        th, td {
            border: 1px solid #333;
            padding: 2pt 5pt;
            text-align: left;
            font-size: 10pt;
            line-height: 1.3;
        }
        th {
            font-weight: bold;
        }
        .pname { width: 45%; }
        .pval { width: 55%; }
        .result-line {
            margin: 2pt 0;
            font-size: 10.5pt;
            line-height: 1.3;
        }
        .step-title {
            font-weight: bold;
            font-size: 11pt;
            margin-top: 6pt;
            margin-bottom: 2pt;
        }
        .calc-item {
            margin: 1pt 0;
            font-size: 10pt;
        }
        .formula { font-weight: bold; }
        .conclusion {
            background: #fafafa;
            padding: 4pt 8pt;
            border-left: 1.5pt solid #333;
            margin: 4pt 0;
            font-size: 10.5pt;
            line-height: 1.4;
        }
        .footer {
            text-align: right;
            margin-top: 10pt;
            font-size: 9.5pt;
        }
    </style>
</head>
<body>
    <div class="report-title">胶粘剂性能评估报告</div>

    <!-- ========== 1. 项目信息 ========== -->
    <div class="section-title">${inputParams[0].title}</div>
`;
        for (const [label, val] of inputParams[0].items) {
            html += '    <div class="info-item">' + label + '  ' + val + '</div>\n';
        }

        // ========== 2~6. 输入参数（分组表格） ==========
        html += '\n    <div class="section-title">输入参数</div>\n';
        html += paramTablesHTML;

        // ========== 关键参数摘要 ==========
        html += '\n    <div class="section-title">关键参数摘要</div>\n';
        html += '<table>\n';
        html += '<tr><th class="pname">参数名称</th><th class="pval">数值</th></tr>\n';
        html += '<tr><td class="pname">玻璃热膨胀系数</td><td class="pval">' + inputs.glass_thermal_exp + ' /℃</td></tr>\n';
        html += '<tr><td class="pname">窗框热膨胀系数</td><td class="pval">' + inputs.frame_thermal_exp + ' /℃</td></tr>\n';
        html += '<tr><td class="pname">玻璃钢热膨胀系数</td><td class="pval">' + inputs.frp_thermal_exp + ' /℃</td></tr>\n';
        html += '<tr><td class="pname">胶层面积</td><td class="pval">' + fmt(results.bond_area, 2) + ' mm²</td></tr>\n';
        html += '<tr><td class="pname">剪切应力</td><td class="pval">' + fmt(results.tau_shear, 6) + ' MPa</td></tr>\n';
        html += '<tr><td class="pname">拉伸应力</td><td class="pval">' + fmt(results.sigma_tensile, 6) + ' MPa</td></tr>\n';
        html += '<tr><td class="pname">总老化系数</td><td class="pval">' + fmt(results.total_aging_factor, 6) + '</td></tr>\n';
        html += '</table>\n';

        // ========== 评估结果 ==========
        html += '\n    <div class="section-title">评估结果</div>\n';
        html += '<div class="result-line">总体评估: ' + (results.is_safe ? '胶粘剂满足要求' : '胶粘剂不满足要求') + '</div>\n';
        html += '<div class="result-line">剪切安全系数: ' + fmt(results.shear_safety_factor, 2) + ' (要求: > ' + inputs.shear_req + ')</div>\n';
        html += '<div class="result-line">拉伸安全系数: ' + fmt(results.tensile_safety_factor, 2) + ' (要求: > ' + inputs.tensile_req + ')</div>\n';
        html += '<div class="result-line">剪切变形角: ' + fmt(results.tan_theta, 6) + ' (要求: ≤ ' + inputs.allow_shear_angle + ')</div>\n';

        // ========== 隧道及会车工况 ==========
        html += '\n    <div class="section-title">隧道及会车工况安全系数</div>\n';
        html += conditionHTML;

        // ========== 详细计算过程 ==========
        html += '\n    <div class="section-title">详细计算过程</div>\n';
        html += stepsHTML;

        // ========== 结论与建议 + 编审批 ==========
        html += '\n    <div class="section-title">结论与建议</div>\n';
        html += '    <div class="conclusion">' +
            (results.is_safe
                ? '根据计算结果，该胶粘剂的各项性能指标均满足设计要求，包括隧道和会车工况，可以在本项目中使用。'
                : '根据计算结果，该胶粘剂部分性能指标不满足设计要求，建议调整胶粘剂参数或更换更适合的胶粘剂类型。') +
            '</div>\n';

        html += '\n    <div class="section-title">编审批</div>\n';
        html += '<table>\n';
        html += '<tr><th style="width:30%"></th><th style="width:35%">姓名</th><th style="width:35%">日期</th></tr>\n';
        html += '<tr><td>编制</td><td></td><td></td></tr>\n';
        html += '<tr><td>审核</td><td></td><td></td></tr>\n';
        html += '<tr><td>批准</td><td></td><td></td></tr>\n';
        html += '</table>\n';

        html += '\n    <div class="footer">报告生成日期: ' + dateStr + '</div>\n';

        html += '\n    <script>\n        window.onload = function() { setTimeout(function() { window.print(); }, 500); };\n    </script>\n';
        html += '</body>\n</html>';

        const win = window.open('', '_blank');
        win.document.open();
        win.document.write(html);
        win.document.close();

    } catch (e) {
        console.error('导出PDF失败:', e);
        alert('导出PDF失败: ' + (e.message || e));
    }
}
