// ========== 报告导出功能 ==========
// Word: 使用 docx.js 生成真实的 .docx（OOXML 格式）
// PDF:  使用 HTML + window.print() 导出，排版与 Word 完全一致
// 注意：docx.umd.js 和 FileSaver.min.js 需要在 index.html 中先加载

// 工具函数：格式化数字
function fmt(num, digits) {
    if (num === undefined || num === null || isNaN(num)) return '0';
    return Number(num).toFixed(digits);
}

// ========== 导出 Word 报告 ==========
async function exportWordReport() {
    if (!window.calculationResults) {
        alert('请先点击"开始计算"！');
        return;
    }

    try {
        const inputs = window.calculationResults.inputs;
        const results = window.calculationResults.results;
        const steps = window.calculationResults.steps;

        const D = window.docx;
        if (!D) {
            alert('docx 库未加载，请检查网络或刷新页面重试。');
            return;
        }

        const children = [];

        // ===== 标题 =====
        children.push(new D.Paragraph({
            children: [new D.TextRun({ text: '胶粘剂性能评估报告', bold: true, size: 28, font: '宋体' })],
            alignment: D.AlignmentType.CENTER,
            spacing: { after: 200 }
        }));

        // ===== 项目信息 =====
        children.push(new D.Paragraph({
            children: [new D.TextRun({ text: '项目信息', bold: true, size: 24, font: '宋体' })],
            spacing: { before: 200, after: 120 }
        }));

        // 项目信息表格（3 行 2 列，跟桌面版一致）
        const projectRows = [
            ['项目名称', inputs.project_name || '（未填写）'],
            ['产品图号', inputs.drawing_number || '（未填写）'],
            ['胶粘剂型号', inputs.adhesive_model || '（未填写）']
        ];
        children.push(new D.Table({
            width: { size: 100, type: D.WidthType.PERCENTAGE },
            rows: projectRows.map((row, idx) =>
                new D.TableRow({
                    children: row.map((cellText, colIdx) =>
                        new D.TableCell({
                            width: { size: colIdx === 0 ? 40 : 60, type: D.WidthType.PERCENTAGE },
                            children: [new D.Paragraph({
                                children: [new D.TextRun({
                                    text: cellText,
                                    bold: idx === 0,
                                    size: 20,
                                    font: '宋体'
                                })]
                            })]
                        })
                    )
                })
            )
        }));

        // ===== 输入参数（分组表格，跟桌面版一致）=====
        children.push(new D.Paragraph({
            children: [new D.TextRun({ text: '输入参数', bold: true, size: 24, font: '宋体' })],
            spacing: { before: 300, after: 120 }
        }));

        // 2. 玻璃及安装参数
        const glassData = [
            ['玻璃质量', fmt(inputs.glass_mass, 2) + ' kg'],
            ['玻璃面积', fmt(inputs.glass_area, 2) + ' m²'],
            ['玻璃长度', fmt(inputs.glass_length, 0) + ' mm'],
            ['玻璃宽度', fmt(inputs.glass_width, 0) + ' mm'],
            ['安装角度', fmt(inputs.installation_angle, 0) + ' °'],
            ['粘接宽度', fmt(inputs.bond_width, 0) + ' mm'],
            ['胶层厚度', fmt(inputs.adhesive_thickness, 2) + ' mm'],
            ['列车最大速度', fmt(inputs.train_speed, 0) + ' km/h']
        ];
        children.push(new D.Paragraph({
            children: [new D.TextRun({ text: '2. 玻璃及安装参数', bold: true, size: 22, font: '宋体' })],
            spacing: { before: 200, after: 80 }
        }));
        children.push(buildParamTable(D, glassData));

        // 3. 环境与载荷参数
        const envData = [
            ['风压', fmt(inputs.wind_pressure, 2) + ' Pa'],
            ['纵向动载荷加速度', fmt(inputs.long_acceleration, 2) + ' g'],
            ['横向动载荷加速度', fmt(inputs.trans_acceleration, 2) + ' g'],
            ['垂直动载荷加速度', fmt(inputs.vert_acceleration, 2) + ' g'],
            ['最低工作温度', fmt(inputs.min_temperature, 2) + ' ℃'],
            ['最高工作温度', fmt(inputs.max_temperature, 2) + ' ℃']
        ];
        children.push(new D.Paragraph({
            children: [new D.TextRun({ text: '3. 环境与载荷参数', bold: true, size: 22, font: '宋体' })],
            spacing: { before: 200, after: 80 }
        }));
        children.push(buildParamTable(D, envData));

        // 4. 材料性能参数
        const matData = [
            ['玻璃热膨胀系数', fmt(inputs.glass_expansion, 2) + ' /℃'],
            ['窗框热膨胀系数', fmt(inputs.frame_expansion, 2) + ' /℃'],
            ['玻璃钢热膨胀系数', fmt(inputs.frp_expansion, 2) + ' /℃'],
            ['老化后拉伸强度', fmt(inputs.tensile_strength, 2) + ' MPa'],
            ['老化后剪切强度', fmt(inputs.shear_strength, 2) + ' MPa'],
            ['允许最大剪切变形角', fmt(inputs.allow_shear_angle, 2)]
        ];
        children.push(new D.Paragraph({
            children: [new D.TextRun({ text: '4. 材料性能参数', bold: true, size: 22, font: '宋体' })],
            spacing: { before: 200, after: 80 }
        }));
        children.push(buildParamTable(D, matData));

        // 5. 老化系数参数
        const agingData = [
            ['温度差异系数', fmt(inputs.temp_factor, 2)],
            ['介质影响系数', fmt(inputs.media_factor, 2)],
            ['几何差异系数', fmt(inputs.geo_factor, 2)],
            ['长期静态系数', fmt(inputs.long_static_factor, 2)],
            ['动态载荷系数', fmt(inputs.dynamic_factor, 2)],
            ['人工操作系数', fmt(inputs.manual_factor, 2)]
        ];
        children.push(new D.Paragraph({
            children: [new D.TextRun({ text: '5. 老化系数参数', bold: true, size: 22, font: '宋体' })],
            spacing: { before: 200, after: 80 }
        }));
        children.push(buildParamTable(D, agingData));

        // 6. 安全系数要求
        const sfData = [
            ['剪切安全系数要求', '> ' + fmt(inputs.shear_safety_req, 1)],
            ['拉伸安全系数要求', '> ' + fmt(inputs.tensile_safety_req, 1)],
            ['户外高速工况', '> ' + fmt(inputs.front_outdoor_req, 1)],
            ['隧道入口工况', '> ' + fmt(inputs.front_tunnel_entry_req, 1)],
            ['隧道内稳定工况', '> ' + fmt(inputs.front_tunnel_stable_req, 1)],
            ['隧道出口工况', '> ' + fmt(inputs.front_tunnel_exit_req, 1)],
            ['会车叠加工况', '> ' + fmt(inputs.front_overtaking_req, 1)]
        ];
        children.push(new D.Paragraph({
            children: [new D.TextRun({ text: '6. 安全系数要求', bold: true, size: 22, font: '宋体' })],
            spacing: { before: 200, after: 80 }
        }));
        children.push(buildParamTable(D, sfData));

        // ===== 关键参数摘要 =====
        children.push(new D.Paragraph({
            children: [new D.TextRun({ text: '关键参数摘要', bold: true, size: 24, font: '宋体' })],
            spacing: { before: 300, after: 120 }
        }));

        const summaryData = [
            ['胶层面积', fmt(results.bond_area, 2) + ' mm²'],
            ['剪切应力', fmt(results.tau_shear, 6) + ' MPa'],
            ['拉伸应力', fmt(results.sigma_tensile, 6) + ' MPa'],
            ['剪切变形角', fmt(results.tan_theta, 6)],
            ['总老化系数', fmt(results.total_aging_factor, 6)],
            ['老化后剪切强度', fmt(results.aged_shear_strength, 2) + ' MPa'],
            ['老化后拉伸强度', fmt(results.aged_tensile_strength, 2) + ' MPa'],
            ['剪切安全系数', fmt(results.shear_safety_factor, 2)],
            ['拉伸安全系数', fmt(results.tensile_safety_factor, 2)]
        ];
        children.push(buildParamTable(D, summaryData));

        // ===== 评估结果 =====
        children.push(new D.Paragraph({
            children: [new D.TextRun({ text: '评估结果', bold: true, size: 24, font: '宋体' })],
            spacing: { before: 300, after: 120 }
        }));

        children.push(new D.Paragraph({
            children: [new D.TextRun({
                text: '总体评估: ' + (results.is_safe ? '胶粘剂满足要求' : '胶粘剂不满足要求'),
                bold: true, size: 22, font: '宋体',
                color: results.is_safe ? '008000' : 'FF0000'
            })],
            spacing: { after: 80 }
        }));

        children.push(new D.Paragraph({
            children: [new D.TextRun({
                text: '剪切安全系数: ' + fmt(results.shear_safety_factor, 2) + ' (要求: > ' + inputs.shear_req + ')  → ' + (results.is_shear_safe ? '满足' : '不满足'),
                size: 22, font: '宋体',
                color: results.is_shear_safe ? '008000' : 'FF0000'
            })],
            spacing: { after: 80 }
        }));

        children.push(new D.Paragraph({
            children: [new D.TextRun({
                text: '拉伸安全系数: ' + fmt(results.tensile_safety_factor, 2) + ' (要求: > ' + inputs.tensile_req + ')  → ' + (results.is_tensile_safe ? '满足' : '不满足'),
                size: 22, font: '宋体',
                color: results.is_tensile_safe ? '008000' : 'FF0000'
            })],
            spacing: { after: 80 }
        }));

        children.push(new D.Paragraph({
            children: [new D.TextRun({
                text: '剪切变形角: ' + fmt(results.tan_theta, 6) + ' (要求: ≤ ' + inputs.allow_shear_angle + ')  → ' + (results.is_angle_safe ? '满足' : '不满足'),
                size: 22, font: '宋体',
                color: results.is_angle_safe ? '008000' : 'FF0000'
            })],
            spacing: { after: 80 }
        }));

        // ===== 隧道及会车工况安全系数 =====
        if (results.condition_safety_factors && Object.keys(results.condition_safety_factors).length > 0) {
            children.push(new D.Paragraph({
                children: [new D.TextRun({ text: '隧道及会车工况安全系数', bold: true, size: 24, font: '宋体' })],
                spacing: { before: 300, after: 120 }
            }));

            const conditionNames = {
                'outdoor': '户外高速',
                'tunnel_entry': '隧道入口',
                'tunnel_stable': '隧道内稳定',
                'tunnel_exit': '隧道出口',
                'overtaking': '会车叠加'
            };

            for (const [condition, factor] of Object.entries(results.condition_safety_factors)) {
                const req = results.condition_requirements ? results.condition_requirements[condition] : 1.5;
                const passed = results.condition_safety_status ? results.condition_safety_status[condition] : factor > req;
                const name = conditionNames[condition] || condition;

                children.push(new D.Paragraph({
                    children: [new D.TextRun({
                        text: name + ': ' + fmt(factor, 2) + ' (要求: > ' + fmt(req, 1) + ')  → ' + (passed ? '满足' : '不满足'),
                        size: 22, font: '宋体',
                        color: passed ? '008000' : 'FF0000'
                    })],
                    spacing: { after: 80 }
                }));
            }
        }

        // ===== 详细计算过程 =====
        children.push(new D.Paragraph({
            children: [new D.TextRun({ text: '详细计算过程', bold: true, size: 24, font: '宋体' })],
            spacing: { before: 400, after: 120 }
        }));

        for (const step of steps) {
            children.push(new D.Paragraph({
                children: [new D.TextRun({ text: step.title, bold: true, size: 22, font: '宋体' })],
                spacing: { before: 200, after: 100 }
            }));

            for (const calc of step.calculations) {
                children.push(new D.Paragraph({
                    children: [new D.TextRun({ text: '公式: ' + calc.formula, bold: true, size: 20, font: '宋体' })],
                    spacing: { after: 60 }
                }));
                children.push(new D.Paragraph({
                    children: [new D.TextRun({ text: '计算: ' + calc.value, size: 20, font: '宋体' })],
                    spacing: { after: 80 }
                }));
            }
        }

        // ===== 结论与建议 =====
        children.push(new D.Paragraph({
            children: [new D.TextRun({ text: '结论与建议', bold: true, size: 24, font: '宋体' })],
            spacing: { before: 400, after: 120 }
        }));

        const conclusion = results.is_safe
            ? '根据计算结果，该胶粘剂的各项性能指标均满足设计要求，包括隧道和会车工况，可以在本项目中使用。'
            : '根据计算结果，该胶粘剂部分性能指标不满足设计要求，建议调整胶粘剂参数、更换更适合的胶粘剂类型，或重新校核载荷条件。';

        children.push(new D.Paragraph({
            children: [new D.TextRun({ text: conclusion, size: 22, font: '宋体' })],
            spacing: { after: 100 }
        }));

        // ===== 编审批 =====
        children.push(new D.Paragraph({
            children: [new D.TextRun({ text: '编审批', bold: true, size: 24, font: '宋体' })],
            spacing: { before: 400, after: 120 }
        }));

        // 编审批表格（4 行 3 列）
        const approvalRows = [
            ['', '姓名', '日期'],
            ['编制', '', ''],
            ['审核', '', ''],
            ['批准', '', '']
        ];
        children.push(new D.Table({
            width: { size: 100, type: D.WidthType.PERCENTAGE },
            rows: approvalRows.map((row, idx) =>
                new D.TableRow({
                    children: row.map((cellText, colIdx) =>
                        new D.TableCell({
                            width: { size: colIdx === 0 ? 30 : 35, type: D.WidthType.PERCENTAGE },
                            children: [new D.Paragraph({
                                children: [new D.TextRun({
                                    text: cellText,
                                    bold: idx === 0,
                                    size: 22, font: '宋体'
                                })]
                            })]
                        })
                    )
                })
            )
        }));

        // 报告日期
        const now = new Date();
        const dateStr = now.getFullYear() + '-' +
            String(now.getMonth() + 1).padStart(2, '0') + '-' +
            String(now.getDate()).padStart(2, '0') + ' ' +
            String(now.getHours()).padStart(2, '0') + ':' +
            String(now.getMinutes()).padStart(2, '0');

        children.push(new D.Paragraph({
            children: [new D.TextRun({ text: '报告生成日期: ' + dateStr, size: 20, font: '宋体' })],
            spacing: { before: 300 },
            alignment: D.AlignmentType.RIGHT
        }));

        // ===== 创建 Word 文档（A4 页面，边距默认适中）=====
        const doc = new D.Document({
            creator: '胶粘剂性能评估系统',
            title: '',
            sections: [{
                properties: {
                    page: {
                        size: { width: 11906, height: 16838 },
                        margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 }
                    }
                },
                children: children
            }]
        });

        // ===== 生成并下载 =====
        const blob = await D.Packer.toBlob(doc);
        const filename = '胶粘剂性能评估报告_' +
            now.getFullYear() +
            String(now.getMonth() + 1).padStart(2, '0') +
            String(now.getDate()).padStart(2, '0') +
            '.docx';

        if (window.saveAs) {
            window.saveAs(blob, filename);
        } else {
            // 如果 FileSaver.js 也没加载，用原生方法
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = filename;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }

    } catch (e) {
        console.error('导出Word失败:', e);
        alert('导出Word失败: ' + (e.message || e));
    }
}

// ========== 辅助函数：生成"参数名称/数值"表格 ==========
function buildParamTable(D, dataRows) {
    const rows = [
        new D.TableRow({
            children: [
                new D.TableCell({
                    width: { size: 45, type: D.WidthType.PERCENTAGE },
                    children: [new D.Paragraph({
                        children: [new D.TextRun({ text: '参数名称', bold: true, size: 20, font: '宋体' })]
                    })]
                }),
                new D.TableCell({
                    width: { size: 55, type: D.WidthType.PERCENTAGE },
                    children: [new D.Paragraph({
                        children: [new D.TextRun({ text: '数值', bold: true, size: 20, font: '宋体' })]
                    })]
                })
            ]
        }),
        ...dataRows.map(row => new D.TableRow({
            children: [
                new D.TableCell({
                    width: { size: 45, type: D.WidthType.PERCENTAGE },
                    children: [new D.Paragraph({ children: [new D.TextRun({ text: row[0], size: 20, font: '宋体' })] })]
                }),
                new D.TableCell({
                    width: { size: 55, type: D.WidthType.PERCENTAGE },
                    children: [new D.Paragraph({ children: [new D.TextRun({ text: row[1], size: 20, font: '宋体' })] })]
                })
            ]
        }))
    ];

    return new D.Table({
        width: { size: 100, type: D.WidthType.PERCENTAGE },
        rows: rows
    });
}

// ========== 导出 PDF 报告（与 Word 排版完全一致）==========
async function exportPdfReport() {
    if (!window.calculationResults) {
        alert('请先点击"开始计算"！');
        return;
    }

    try {
        const inputs = window.calculationResults.inputs;
        const results = window.calculationResults.results;
        const steps = window.calculationResults.steps;

        // ===== 构建参数表格 HTML =====
        function buildHtmlParamTable(dataRows) {
            let html = '<table><tr><th class="pname">参数名称</th><th class="pval">数值</th></tr>';
            for (const row of dataRows) {
                html += '<tr><td class="pname">' + row[0] + '</td><td class="pval">' + row[1] + '</td></tr>';
            }
            html += '</table>';
            return html;
        }

        // ===== 项目信息 =====
        const projectRows = [
            ['项目名称', inputs.project_name || '（未填写）'],
            ['产品图号', inputs.drawing_number || '（未填写）'],
            ['胶粘剂型号', inputs.adhesive_model || '（未填写）']
        ];

        // ===== 输入参数 =====
        const glassData = [
            ['玻璃质量', fmt(inputs.glass_mass, 2) + ' kg'],
            ['玻璃面积', fmt(inputs.glass_area, 2) + ' m²'],
            ['玻璃长度', fmt(inputs.glass_length, 0) + ' mm'],
            ['玻璃宽度', fmt(inputs.glass_width, 0) + ' mm'],
            ['安装角度', fmt(inputs.installation_angle, 0) + ' °'],
            ['粘接宽度', fmt(inputs.bond_width, 0) + ' mm'],
            ['胶层厚度', fmt(inputs.adhesive_thickness, 2) + ' mm'],
            ['列车最大速度', fmt(inputs.train_speed, 0) + ' km/h']
        ];

        const envData = [
            ['风压', fmt(inputs.wind_pressure, 2) + ' Pa'],
            ['纵向动载荷加速度', fmt(inputs.long_acceleration, 2) + ' g'],
            ['横向动载荷加速度', fmt(inputs.trans_acceleration, 2) + ' g'],
            ['垂直动载荷加速度', fmt(inputs.vert_acceleration, 2) + ' g'],
            ['最低工作温度', fmt(inputs.min_temperature, 2) + ' ℃'],
            ['最高工作温度', fmt(inputs.max_temperature, 2) + ' ℃']
        ];

        const matData = [
            ['玻璃热膨胀系数', fmt(inputs.glass_expansion, 2) + ' /℃'],
            ['窗框热膨胀系数', fmt(inputs.frame_expansion, 2) + ' /℃'],
            ['玻璃钢热膨胀系数', fmt(inputs.frp_expansion, 2) + ' /℃'],
            ['老化后拉伸强度', fmt(inputs.tensile_strength, 2) + ' MPa'],
            ['老化后剪切强度', fmt(inputs.shear_strength, 2) + ' MPa'],
            ['允许最大剪切变形角', fmt(inputs.allow_shear_angle, 2)]
        ];

        const agingData = [
            ['温度差异系数', fmt(inputs.temp_factor, 2)],
            ['介质影响系数', fmt(inputs.media_factor, 2)],
            ['几何差异系数', fmt(inputs.geo_factor, 2)],
            ['长期静态系数', fmt(inputs.long_static_factor, 2)],
            ['动态载荷系数', fmt(inputs.dynamic_factor, 2)],
            ['人工操作系数', fmt(inputs.manual_factor, 2)]
        ];

        const sfData = [
            ['剪切安全系数要求', '> ' + fmt(inputs.shear_safety_req, 1)],
            ['拉伸安全系数要求', '> ' + fmt(inputs.tensile_safety_req, 1)],
            ['户外高速工况', '> ' + fmt(inputs.front_outdoor_req, 1)],
            ['隧道入口工况', '> ' + fmt(inputs.front_tunnel_entry_req, 1)],
            ['隧道内稳定工况', '> ' + fmt(inputs.front_tunnel_stable_req, 1)],
            ['隧道出口工况', '> ' + fmt(inputs.front_tunnel_exit_req, 1)],
            ['会车叠加工况', '> ' + fmt(inputs.front_overtaking_req, 1)]
        ];

        const summaryData = [
            ['胶层面积', fmt(results.bond_area, 2) + ' mm²'],
            ['剪切应力', fmt(results.tau_shear, 6) + ' MPa'],
            ['拉伸应力', fmt(results.sigma_tensile, 6) + ' MPa'],
            ['剪切变形角', fmt(results.tan_theta, 6)],
            ['总老化系数', fmt(results.total_aging_factor, 6)],
            ['老化后剪切强度', fmt(results.aged_shear_strength, 2) + ' MPa'],
            ['老化后拉伸强度', fmt(results.aged_tensile_strength, 2) + ' MPa'],
            ['剪切安全系数', fmt(results.shear_safety_factor, 2)],
            ['拉伸安全系数', fmt(results.tensile_safety_factor, 2)]
        ];

        // ===== 评估结果 =====
        function evalResult(ok) { return ok ? '<span style="color:#008000;">✓ 满足</span>' : '<span style="color:#d00;">✗ 不满足</span>'; }

        // 隧道及会车工况
        let conditionHTML = '';
        if (results.condition_safety_factors && Object.keys(results.condition_safety_factors).length > 0) {
            const conditionNames = {
                'outdoor': '户外高速',
                'tunnel_entry': '隧道入口',
                'tunnel_stable': '隧道内稳定',
                'tunnel_exit': '隧道出口',
                'overtaking': '会车叠加'
            };
            for (const [condition, factor] of Object.entries(results.condition_safety_factors)) {
                const req = results.condition_requirements ? results.condition_requirements[condition] : 1.5;
                const passed = results.condition_safety_status ? results.condition_safety_status[condition] : factor > req;
                const name = conditionNames[condition] || condition;
                conditionHTML += '<div class="result-line">' + name + ': ' + fmt(factor, 2) + ' (要求: > ' + fmt(req, 1) + ')  → ' + evalResult(passed) + '</div>';
            }
        }

        // ===== 详细计算过程 =====
        let stepsHTML = '';
        for (const step of steps) {
            stepsHTML += '<div class="step-title">' + step.title + '</div>';
            for (const calc of step.calculations) {
                stepsHTML += '<div class="calc-item"><strong>公式: ' + calc.formula + '</strong></div>';
                stepsHTML += '<div class="calc-item">计算: ' + calc.value + '</div>';
            }
        }

        // ===== 构建完整 HTML 页面 =====
        const now = new Date();
        const dateStr = now.getFullYear() + '-' +
            String(now.getMonth() + 1).padStart(2, '0') + '-' +
            String(now.getDate()).padStart(2, '0') + ' ' +
            String(now.getHours()).padStart(2, '0') + ':' +
            String(now.getMinutes()).padStart(2, '0');

        const conclusion = results.is_safe
            ? '根据计算结果，该胶粘剂的各项性能指标均满足设计要求，包括隧道和会车工况，可以在本项目中使用。'
            : '根据计算结果，该胶粘剂部分性能指标不满足设计要求，建议调整胶粘剂参数、更换更适合的胶粘剂类型，或重新校核载荷条件。';

        const html = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>胶粘剂性能评估报告</title>
    <style>
        @page {
            size: A4;
            margin: 2.54cm;
        }
        body {
            font-family: "SimSun", "宋体", serif;
            line-height: 1.5;
            color: #000;
            font-size: 11pt;
            margin: 0;
            padding: 10px 20px;
        }
        .toolbar {
            display: flex;
            gap: 10px;
            justify-content: center;
            margin-bottom: 20px;
            padding-bottom: 15px;
            border-bottom: 1px dashed #999;
        }
        .toolbar button {
            font-size: 16px;
            padding: 12px 24px;
            border: 1px solid #2563eb;
            background: #2563eb;
            color: white;
            border-radius: 8px;
            cursor: pointer;
            font-family: inherit;
        }
        .toolbar .btn-back {
            background: #fff;
            color: #2563eb;
        }
        .report-title {
            text-align: center;
            font-size: 16pt;
            font-weight: bold;
            margin: 0 0 18pt 0;
        }
        .section-title {
            font-size: 12pt;
            font-weight: bold;
            margin-top: 16pt;
            margin-bottom: 6pt;
            padding-bottom: 2pt;
            border-bottom: 1pt solid #555;
        }
        .group-title {
            font-size: 11pt;
            font-weight: bold;
            margin-top: 10pt;
            margin-bottom: 4pt;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            margin: 4pt 0 6pt 0;
        }
        th, td {
            border: 1px solid #333;
            padding: 3pt 6pt;
            text-align: left;
            font-size: 10pt;
            line-height: 1.3;
        }
        th {
            font-weight: bold;
            background: #f7f7f7;
        }
        .pname { width: 45%; }
        .pval { width: 55%; }
        .result-line {
            margin: 3pt 0;
            font-size: 11pt;
            line-height: 1.4;
        }
        .step-title {
            font-weight: bold;
            font-size: 11pt;
            margin-top: 10pt;
            margin-bottom: 4pt;
        }
        .calc-item {
            margin: 2pt 0;
            font-size: 10pt;
        }
        .conclusion {
            background: #fafafa;
            padding: 8pt 12pt;
            border-left: 2pt solid #333;
            margin: 6pt 0;
            font-size: 11pt;
            line-height: 1.5;
        }
        .footer {
            text-align: right;
            margin-top: 15pt;
            font-size: 9.5pt;
            color: #333;
        }
        @media print {
            .toolbar { display: none !important; }
            body { padding: 0; }
        }
    </style>
</head>
<body>
    <div class="toolbar">
        <button onclick="window.location.href='index.html'" class="btn-back">← 返回</button>
        <button onclick="window.print()">🖨️ 打印 / 导出PDF</button>
    </div>

    <div class="report-title">胶粘剂性能评估报告</div>

    <div class="section-title">项目信息</div>
    ` + buildHtmlParamTable(projectRows) + `

    <div class="section-title">输入参数</div>

    <div class="group-title">2. 玻璃及安装参数</div>
    ` + buildHtmlParamTable(glassData) + `

    <div class="group-title">3. 环境与载荷参数</div>
    ` + buildHtmlParamTable(envData) + `

    <div class="group-title">4. 材料性能参数</div>
    ` + buildHtmlParamTable(matData) + `

    <div class="group-title">5. 老化系数参数</div>
    ` + buildHtmlParamTable(agingData) + `

    <div class="group-title">6. 安全系数要求</div>
    ` + buildHtmlParamTable(sfData) + `

    <div class="section-title">关键参数摘要</div>
    ` + buildHtmlParamTable(summaryData) + `

    <div class="section-title">评估结果</div>
    <div class="result-line"><strong>总体评估: ` + (results.is_safe ? '胶粘剂满足要求' : '胶粘剂不满足要求') + `</strong></div>
    <div class="result-line">剪切安全系数: ` + fmt(results.shear_safety_factor, 2) + ` (要求: > ` + inputs.shear_req + `)  → ` + evalResult(results.is_shear_safe) + `</div>
    <div class="result-line">拉伸安全系数: ` + fmt(results.tensile_safety_factor, 2) + ` (要求: > ` + inputs.tensile_req + `)  → ` + evalResult(results.is_tensile_safe) + `</div>
    <div class="result-line">剪切变形角: ` + fmt(results.tan_theta, 6) + ` (要求: ≤ ` + inputs.allow_shear_angle + `)  → ` + evalResult(results.is_angle_safe) + `</div>

    <div class="section-title">隧道及会车工况安全系数</div>
    ` + conditionHTML + `

    <div class="section-title">详细计算过程</div>
    ` + stepsHTML + `

    <div class="section-title">结论与建议</div>
    <div class="conclusion">` + conclusion + `</div>

    <div class="section-title">编审批</div>
    <table>
        <tr><th style="width:30%"></th><th style="width:35%">姓名</th><th style="width:35%">日期</th></tr>
        <tr><td>编制</td><td></td><td></td></tr>
        <tr><td>审核</td><td></td><td></td></tr>
        <tr><td>批准</td><td></td><td></td></tr>
    </table>

    <div class="footer">报告生成日期: ` + dateStr + `</div>
</body>
</html>`;

        // ===== 替换当前页面（避免手机端开新标签页迷路）=====
        document.open();
        document.write(html);
        document.close();
        document.title = '胶粘剂性能评估报告';

    } catch (e) {
        console.error('导出PDF失败:', e);
        alert('导出PDF失败: ' + (e.message || e));
    }
}
