const CONDITION_NAMES = {
    'outdoor': '户外高速',
    'tunnel_entry': '隧道入口',
    'tunnel_stable': '隧道内稳定',
    'tunnel_exit': '隧道出口',
    'overtaking': '会车叠加'
};

document.addEventListener('DOMContentLoaded', () => {
    setupEventListeners();
});

function setupEventListeners() {
    // Tab切换
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', switchTab);
    });

    // 按钮事件
    document.getElementById('calculate-btn').addEventListener('click', calculate);
    document.getElementById('reset-btn').addEventListener('click', resetParameters);
    document.getElementById('back-to-input-btn').addEventListener('click', backToInput);
    document.getElementById('export-word-btn').addEventListener('click', exportWordReport);
    document.getElementById('export-pdf-btn').addEventListener('click', exportPdfReport);
}

function switchTab(event) {
    const targetTab = event.target.dataset.tab;
    showTab(targetTab);
}

function showTab(tabName) {
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');

    document.querySelectorAll('.tab-panel').forEach(content => {
        content.style.display = 'none';
    });
    document.getElementById(tabName).style.display = 'block';
}

function backToInput() {
    showTab('input');
    window.scrollTo(0, 0);
}

function collectParameters() {
    return {
        project_name: document.getElementById('project_name').value || '未命名项目',
        drawing_number: document.getElementById('drawing_number').value || '无图号',
        adhesive_model: document.getElementById('adhesive_model').value || '未指定',
        window_position: document.getElementById('window_position').value,
        gravity: 9.81,

        glass_mass: parseFloat(document.getElementById('glass_mass').value),
        glass_area: parseFloat(document.getElementById('glass_area').value),
        glass_length: parseFloat(document.getElementById('glass_length').value),
        glass_width: parseFloat(document.getElementById('glass_width').value),
        installation_angle: parseFloat(document.getElementById('installation_angle').value),
        bond_width: parseFloat(document.getElementById('bond_width').value),
        adhesive_thickness: parseFloat(document.getElementById('adhesive_thickness').value),
        train_speed: parseFloat(document.getElementById('train_speed').value),

        wind_pressure: parseFloat(document.getElementById('wind_pressure').value),
        long_acceleration: parseFloat(document.getElementById('long_acceleration').value),
        trans_acceleration: parseFloat(document.getElementById('trans_acceleration').value),
        vert_acceleration: parseFloat(document.getElementById('vert_acceleration').value),
        min_temperature: parseFloat(document.getElementById('min_temperature').value),
        max_temperature: parseFloat(document.getElementById('max_temperature').value),

        glass_thermal_exp: parseFloat(document.getElementById('glass_expansion').value),
        frame_thermal_exp: parseFloat(document.getElementById('frame_expansion').value),
        frp_thermal_exp: parseFloat(document.getElementById('frp_expansion').value),
        tensile_strength: parseFloat(document.getElementById('tensile_strength').value),
        shear_strength: parseFloat(document.getElementById('shear_strength').value),
        allow_shear_angle: parseFloat(document.getElementById('allow_shear_angle').value),

        temp_factor: parseFloat(document.getElementById('temp_factor').value),
        media_factor: parseFloat(document.getElementById('medium_factor').value),
        geom_factor: parseFloat(document.getElementById('geo_factor').value),
        long_static_factor: parseFloat(document.getElementById('long_static_factor').value),
        dynamic_factor: parseFloat(document.getElementById('dynamic_factor').value),
        manual_factor: parseFloat(document.getElementById('human_factor').value),

        shear_req: parseFloat(document.getElementById('shear_safety_req').value),
        tensile_req: parseFloat(document.getElementById('tensile_safety_req').value),

        outdoor_highway_req: parseFloat(document.getElementById('front_outdoor_req').value),
        tunnel_entry_req: parseFloat(document.getElementById('front_tunnel_entry_req').value),
        tunnel_stable_req: parseFloat(document.getElementById('front_tunnel_stable_req').value),
        tunnel_exit_req: parseFloat(document.getElementById('front_tunnel_exit_req').value),
        meeting_req: parseFloat(document.getElementById('front_overtaking_req').value),

        // 根据窗户位置，使用对应工况要求
        is_front_window: document.getElementById('window_position').value === '前窗'
    };
}

function validateParameters(params) {
    const positiveParams = [
        'glass_mass', 'glass_area', 'glass_length', 'glass_width',
        'bond_width', 'adhesive_thickness', 'wind_pressure',
        'tensile_strength', 'shear_strength', 'allow_shear_angle'
    ];

    for (const param of positiveParams) {
        if (isNaN(params[param]) || params[param] <= 0) {
            alert('参数 "' + param + '" 必须大于0');
            return false;
        }
    }

    if (params.installation_angle < 0 || params.installation_angle > 90) {
        alert('安装角度必须在0-90度之间');
        return false;
    }

    return true;
}

function calculate() {
    try {
        const params = collectParameters();

        if (!validateParameters(params)) {
            return;
        }

        const results = performCalculation(params);

        // 将结果存储到全局对象，供导出报告使用
        window.calculationResults = {
            inputs: params,
            results: results,
            steps: results.steps
        };

        updateResultsDisplay(results, params);

        // 自动切换到结果页
        showTab('results');

    } catch (error) {
        alert('计算错误: ' + error.message);
    }
}

function performCalculation(params) {
    const results = {};
    const steps = [];

    const angleRad = params.installation_angle * Math.PI / 180;

    // 1. 静载荷分解
    const step1 = { title: '1. 静载荷分解', calculations: [] };
    results.f_weight = params.glass_mass * params.gravity;
    step1.calculations.push({
        formula: '玻璃自重 = 玻璃质量 × 重力加速度',
        value: 'F_weight = ' + params.glass_mass + ' kg × ' + params.gravity + ' m/s² = ' + results.f_weight.toFixed(2) + ' N'
    });

    results.f_shear_static = results.f_weight * Math.cos(angleRad);
    step1.calculations.push({
        formula: '切向静载荷 = 玻璃自重 × cos(安装角度)',
        value: 'F_切(静) = ' + results.f_weight.toFixed(2) + ' N × cos(' + params.installation_angle + '°) = ' + results.f_shear_static.toFixed(2) + ' N'
    });

    results.f_normal_static = results.f_weight * Math.sin(angleRad);
    step1.calculations.push({
        formula: '垂向静载荷 = 玻璃自重 × sin(安装角度)',
        value: 'F_垂(静) = ' + results.f_weight.toFixed(2) + ' N × sin(' + params.installation_angle + '°) = ' + results.f_normal_static.toFixed(2) + ' N'
    });
    steps.push(step1);

    // 2. 动载荷分解
    const step2 = { title: '2. 动载荷分解', calculations: [] };
    results.a_shear_vert = params.vert_acceleration * params.gravity * Math.cos(angleRad);
    step2.calculations.push({
        formula: '垂向加速度的切向分量 = 垂直动载荷加速度 × g × cos(安装角度)',
        value: 'a_切(垂) = ' + params.vert_acceleration + 'g × ' + params.gravity + ' m/s² × cos(' + params.installation_angle + '°) = ' + results.a_shear_vert.toFixed(2) + ' m/s²'
    });

    results.a_normal_vert = params.vert_acceleration * params.gravity * Math.sin(angleRad);
    step2.calculations.push({
        formula: '垂向加速度的垂向分量 = 垂直动载荷加速度 × g × sin(安装角度)',
        value: 'a_垂(垂) = ' + params.vert_acceleration + 'g × ' + params.gravity + ' m/s² × sin(' + params.installation_angle + '°) = ' + results.a_normal_vert.toFixed(2) + ' m/s²'
    });

    results.a_shear_trans = params.trans_acceleration * params.gravity * Math.sin(angleRad);
    step2.calculations.push({
        formula: '横向加速度的切向分量 = 横向动载荷加速度 × g × sin(安装角度)',
        value: 'a_切(横) = ' + params.trans_acceleration + 'g × ' + params.gravity + ' m/s² × sin(' + params.installation_angle + '°) = ' + results.a_shear_trans.toFixed(2) + ' m/s²'
    });

    results.a_normal_trans = params.trans_acceleration * params.gravity * Math.cos(angleRad);
    step2.calculations.push({
        formula: '横向加速度的垂向分量 = 横向动载荷加速度 × g × cos(安装角度)',
        value: 'a_垂(横) = ' + params.trans_acceleration + 'g × ' + params.gravity + ' m/s² × cos(' + params.installation_angle + '°) = ' + results.a_normal_trans.toFixed(2) + ' m/s²'
    });

    results.a_shear_long = params.long_acceleration * params.gravity * Math.sin(angleRad);
    step2.calculations.push({
        formula: '纵向加速度的切向分量 = 纵向动载荷加速度 × g × sin(安装角度)',
        value: 'a_切(纵) = ' + params.long_acceleration + 'g × ' + params.gravity + ' m/s² × sin(' + params.installation_angle + '°) = ' + results.a_shear_long.toFixed(2) + ' m/s²'
    });

    results.a_normal_long = params.long_acceleration * params.gravity * Math.cos(angleRad);
    step2.calculations.push({
        formula: '纵向加速度的垂向分量 = 纵向动载荷加速度 × g × cos(安装角度)',
        value: 'a_垂(纵) = ' + params.long_acceleration + 'g × ' + params.gravity + ' m/s² × cos(' + params.installation_angle + '°) = ' + results.a_normal_long.toFixed(2) + ' m/s²'
    });

    results.max_shear_accel_long = Math.abs(results.a_shear_vert) + Math.abs(results.a_shear_long);
    step2.calculations.push({
        formula: '最大纵向切向加速度 = |垂向加速度的切向分量| + |纵向加速度的切向分量|',
        value: 'a_切(动纵max) = |' + results.a_shear_vert.toFixed(2) + '| + |' + results.a_shear_long.toFixed(2) + '| = ' + results.max_shear_accel_long.toFixed(2) + ' m/s²'
    });

    results.max_normal_accel = Math.abs(results.a_normal_vert) + Math.abs(results.a_normal_long) + Math.abs(results.a_normal_trans);
    step2.calculations.push({
        formula: '最大垂向加速度 = |垂向加速度的垂向分量| + |纵向加速度的垂向分量| + |横向加速度的垂向分量|',
        value: 'a_垂(动max) = |' + results.a_normal_vert.toFixed(2) + '| + |' + results.a_normal_long.toFixed(2) + '| + |' + results.a_normal_trans.toFixed(2) + '| = ' + results.max_normal_accel.toFixed(2) + ' m/s²'
    });
    steps.push(step2);

    // 3. 动载荷计算
    const step3 = { title: '3. 动载荷计算', calculations: [] };
    results.f_shear_dyn_long = params.glass_mass * results.max_shear_accel_long;
    step3.calculations.push({
        formula: '纵向切向动载荷 = 玻璃质量 × 最大纵向切向加速度',
        value: 'F_切(动纵) = ' + params.glass_mass + ' kg × ' + results.max_shear_accel_long.toFixed(2) + ' m/s² = ' + results.f_shear_dyn_long.toFixed(2) + ' N'
    });

    results.f_shear_dyn_trans = params.glass_mass * results.a_shear_trans;
    step3.calculations.push({
        formula: '横向切向动载荷 = 玻璃质量 × 最大横向切向加速度',
        value: 'F_切(动横) = ' + params.glass_mass + ' kg × ' + results.a_shear_trans.toFixed(2) + ' m/s² = ' + results.f_shear_dyn_trans.toFixed(2) + ' N'
    });

    results.f_normal_dyn = params.glass_mass * results.max_normal_accel;
    step3.calculations.push({
        formula: '垂向动载荷 = 玻璃质量 × 最大垂向加速度',
        value: 'F_垂(动) = ' + params.glass_mass + ' kg × ' + results.max_normal_accel.toFixed(2) + ' m/s² = ' + results.f_normal_dyn.toFixed(2) + ' N'
    });
    steps.push(step3);

    // 4. 总载荷合成
    const step4 = { title: '4. 总载荷合成', calculations: [] };
    results.f_shear_long_max = results.f_shear_dyn_long + results.f_shear_static;
    step4.calculations.push({
        formula: '最大纵向切向载荷 = 纵向切向动载荷 + 切向静载荷',
        value: 'F_切(纵max) = ' + results.f_shear_dyn_long.toFixed(2) + ' + ' + results.f_shear_static.toFixed(2) + ' = ' + results.f_shear_long_max.toFixed(2) + ' N'
    });

    results.f_shear_total = Math.sqrt(Math.pow(results.f_shear_long_max, 2) + Math.pow(results.f_shear_dyn_trans, 2));
    step4.calculations.push({
        formula: '总切向载荷 = √(最大纵向切向载荷² + 最大横向切向载荷²)',
        value: 'F_切总 = √(' + results.f_shear_long_max.toFixed(2) + '² + ' + results.f_shear_dyn_trans.toFixed(2) + '²) = ' + results.f_shear_total.toFixed(2) + ' N'
    });

    results.f_wind = params.wind_pressure * params.glass_area;
    step4.calculations.push({
        formula: '风载荷 = 风压 × 玻璃面积',
        value: 'F_风压 = ' + params.wind_pressure + ' Pa × ' + params.glass_area + ' m² = ' + results.f_wind.toFixed(2) + ' N'
    });

    results.f_normal_max = results.f_normal_dyn - results.f_normal_static + results.f_wind;
    step4.calculations.push({
        formula: '最大垂向载荷 = 垂向动载荷 - 垂向静载荷 + 风载荷',
        value: 'F_垂(max) = ' + results.f_normal_dyn.toFixed(2) + ' - ' + results.f_normal_static.toFixed(2) + ' + ' + results.f_wind.toFixed(2) + ' = ' + results.f_normal_max.toFixed(2) + ' N'
    });
    steps.push(step4);

    // 5. 胶层应力计算
    const step5 = { title: '5. 胶层应力计算', calculations: [] };
    results.bond_area = 2 * (params.glass_length + params.glass_width) * params.bond_width;
    step5.calculations.push({
        formula: '胶层面积 = 2 × (玻璃长度 + 玻璃宽度) × 粘接宽度',
        value: 'A = 2 × (' + params.glass_length + ' + ' + params.glass_width + ') × ' + params.bond_width + ' = ' + results.bond_area.toFixed(2) + ' mm²'
    });

    results.tau_shear = results.f_shear_total / results.bond_area;
    step5.calculations.push({
        formula: '剪切应力 = 总切向载荷 / 胶层面积',
        value: 'τ_剪切 = ' + results.f_shear_total.toFixed(2) + ' N / ' + results.bond_area.toFixed(2) + ' mm² = ' + results.tau_shear.toFixed(6) + ' MPa'
    });

    results.sigma_tensile = Math.abs(results.f_normal_max) / results.bond_area;
    step5.calculations.push({
        formula: '拉伸应力 = |最大垂向载荷| / 胶层面积',
        value: 'δ_拉伸 = |' + results.f_normal_max.toFixed(2) + '| N / ' + results.bond_area.toFixed(2) + ' mm² = ' + results.sigma_tensile.toFixed(6) + ' MPa'
    });
    steps.push(step5);

    // 6. 胶层热形变计算
    const step6 = { title: '6. 胶层热形变计算', calculations: [] };
    results.delta_t = params.max_temperature - params.min_temperature;
    step6.calculations.push({
        formula: '温度差 = 最高工作温度 - 最低工作温度',
        value: 'ΔT = ' + params.max_temperature + '℃ - ' + params.min_temperature + '℃ = ' + results.delta_t + '℃'
    });

    results.delta_l = (params.glass_length / 2) * (params.frame_thermal_exp - params.glass_thermal_exp) * results.delta_t;
    step6.calculations.push({
        formula: '形变差 = (玻璃长度/2) × (窗框热膨胀系数 - 玻璃热膨胀系数) × 温度差',
        value: 'ΔL = (' + params.glass_length + '/2) × (' + params.frame_thermal_exp + ' - ' + params.glass_thermal_exp + ') × ' + results.delta_t + ' = ' + results.delta_l.toFixed(8) + ' mm'
    });

    results.tan_theta = results.delta_l / params.adhesive_thickness;
    step6.calculations.push({
        formula: '剪切变形角 = 形变差 / 胶层厚度',
        value: 'tanθ = ' + results.delta_l.toFixed(8) + ' mm / ' + params.adhesive_thickness + ' mm = ' + results.tan_theta.toFixed(8)
    });
    steps.push(step6);

    // 7. 老化系数计算
    const step7 = { title: '7. 老化系数计算', calculations: [] };
    results.total_aging_factor = params.temp_factor * params.media_factor * params.geom_factor * params.long_static_factor * params.dynamic_factor * params.manual_factor;
    step7.calculations.push({
        formula: '总老化折减系数 = 温度差异系数 × 介质影响系数 × 几何差异系数 × 长期静态系数 × 动态载荷系数 × 人工操作系数',
        value: 'C_老 = ' + params.temp_factor + ' × ' + params.media_factor + ' × ' + params.geom_factor + ' × ' + params.long_static_factor + ' × ' + params.dynamic_factor + ' × ' + params.manual_factor + ' = ' + results.total_aging_factor.toFixed(6)
    });

    results.aged_tensile_strength = params.tensile_strength * results.total_aging_factor;
    step7.calculations.push({
        formula: '老化后抗拉强度 = 初始抗拉强度 × 总老化折减系数',
        value: 'σ_D拉伸 = ' + params.tensile_strength + ' MPa × ' + results.total_aging_factor.toFixed(6) + ' = ' + results.aged_tensile_strength.toFixed(6) + ' MPa'
    });

    results.aged_shear_strength = params.shear_strength * results.total_aging_factor;
    step7.calculations.push({
        formula: '老化后剪切强度 = 初始剪切强度 × 总老化折减系数',
        value: 'σ_D剪切 = ' + params.shear_strength + ' MPa × ' + results.total_aging_factor.toFixed(6) + ' = ' + results.aged_shear_strength.toFixed(6) + ' MPa'
    });
    steps.push(step7);

    // 8. 安全系数判断
    const step8 = { title: '8. 安全系数判断', calculations: [] };
    results.shear_safety_factor = results.aged_shear_strength / results.tau_shear;
    step8.calculations.push({
        formula: '剪切安全系数 = 老化后剪切强度 / 剪切应力',
        value: '剪切安全系数 = ' + results.aged_shear_strength.toFixed(6) + ' MPa / ' + results.tau_shear.toFixed(6) + ' MPa = ' + results.shear_safety_factor.toFixed(2) + ' (要求: > ' + params.shear_req + ')'
    });

    results.tensile_safety_factor = results.aged_tensile_strength / results.sigma_tensile;
    step8.calculations.push({
        formula: '拉伸安全系数 = 老化后抗拉强度 / 拉伸应力',
        value: '拉伸安全系数 = ' + results.aged_tensile_strength.toFixed(6) + ' MPa / ' + results.sigma_tensile.toFixed(6) + ' MPa = ' + results.tensile_safety_factor.toFixed(2) + ' (要求: > ' + params.tensile_req + ')'
    });
    steps.push(step8);

    // 9. 隧道及会车工况计算
    const step9 = { title: '9. 隧道及会车工况计算', calculations: [] };
    const trainSpeedKmh = params.train_speed || 300;
    const trainSpeed = trainSpeedKmh * 1000 / 3600;
    const airDensity = 1.225;
    const dragCoefficient = params.window_position === '前窗' ? 1.2 : 0.8;

    step9.calculations.push({
        formula: '列车速度转换 = 列车速度 (km/h) × 1000 / 3600',
        value: 'v = ' + trainSpeedKmh + ' km/h × 1000/3600 = ' + trainSpeed.toFixed(2) + ' m/s'
    });

    results.outdoor_pressure = 0.5 * airDensity * Math.pow(trainSpeed, 2) * dragCoefficient;
    step9.calculations.push({
        formula: '户外高速风压 = 0.5 × 空气密度 × 速度² × 风阻系数',
        value: 'P_户外 = 0.5 × ' + airDensity + ' × ' + trainSpeed.toFixed(2) + '² × ' + dragCoefficient + ' = ' + results.outdoor_pressure.toFixed(2) + ' Pa'
    });

    results.tunnel_entry_pressure = results.outdoor_pressure * 1.5;
    step9.calculations.push({
        formula: '隧道入口风压 = 户外高速风压 × 1.5',
        value: 'P_隧道入口 = ' + results.outdoor_pressure.toFixed(2) + ' × 1.5 = ' + results.tunnel_entry_pressure.toFixed(2) + ' Pa'
    });

    results.tunnel_stable_pressure = results.outdoor_pressure * 1.2;
    step9.calculations.push({
        formula: '隧道内稳定风压 = 户外高速风压 × 1.2',
        value: 'P_隧道内 = ' + results.outdoor_pressure.toFixed(2) + ' × 1.2 = ' + results.tunnel_stable_pressure.toFixed(2) + ' Pa'
    });

    results.tunnel_exit_pressure = results.outdoor_pressure * 1.8;
    step9.calculations.push({
        formula: '隧道出口风压 = 户外高速风压 × 1.8',
        value: 'P_隧道出口 = ' + results.outdoor_pressure.toFixed(2) + ' × 1.8 = ' + results.tunnel_exit_pressure.toFixed(2) + ' Pa'
    });

    results.overtaking_pressure = results.outdoor_pressure * 2.0;
    step9.calculations.push({
        formula: '会车风压 = 户外高速风压 × 2.0',
        value: 'P_会车 = ' + results.outdoor_pressure.toFixed(2) + ' × 2.0 = ' + results.overtaking_pressure.toFixed(2) + ' Pa'
    });

    // 确定各工况要求
    const safetyFactorRequirements = params.window_position === '前窗' ? {
        outdoor: params.outdoor_highway_req,
        tunnel_entry: params.tunnel_entry_req,
        tunnel_stable: params.tunnel_stable_req,
        tunnel_exit: params.tunnel_exit_req,
        overtaking: params.meeting_req
    } : {
        outdoor: parseFloat(document.getElementById('side_outdoor_req').value),
        tunnel_entry: parseFloat(document.getElementById('side_tunnel_entry_req').value),
        tunnel_stable: parseFloat(document.getElementById('side_tunnel_stable_req').value),
        tunnel_exit: parseFloat(document.getElementById('side_tunnel_exit_req').value),
        overtaking: parseFloat(document.getElementById('side_overtaking_req').value)
    };

    const conditions = ['outdoor', 'tunnel_entry', 'tunnel_stable', 'tunnel_exit', 'overtaking'];
    const pressureKeys = {
        outdoor: 'outdoor_pressure',
        tunnel_entry: 'tunnel_entry_pressure',
        tunnel_stable: 'tunnel_stable_pressure',
        tunnel_exit: 'tunnel_exit_pressure',
        overtaking: 'overtaking_pressure'
    };

    results.condition_safety_factors = {};
    results.condition_safety_status = {};
    results.condition_requirements = safetyFactorRequirements;

    for (const conditionKey of conditions) {
        const windPressure = results[pressureKeys[conditionKey]];
        const windForce = windPressure * params.glass_area;
        const maxNormalForce = results.f_normal_dyn - results.f_normal_static + windForce;
        const tensileStress = Math.abs(maxNormalForce) / results.bond_area;
        const safetyFactor = results.aged_tensile_strength / tensileStress;

        results.condition_safety_factors[conditionKey] = safetyFactor;
        results.condition_safety_status[conditionKey] = safetyFactor > safetyFactorRequirements[conditionKey];

        const conditionName = CONDITION_NAMES[conditionKey] || conditionKey;
        step9.calculations.push({
            formula: conditionName + '风载荷 = 风压 × 玻璃面积',
            value: 'F_风(' + conditionName + ') = ' + windPressure.toFixed(2) + ' Pa × ' + params.glass_area + ' m² = ' + windForce.toFixed(2) + ' N'
        });
        step9.calculations.push({
            formula: conditionName + '最大垂向载荷 = 垂向动载荷 - 垂向静载荷 + 风载荷',
            value: 'F_垂(max, ' + conditionName + ') = ' + results.f_normal_dyn.toFixed(2) + ' - ' + results.f_normal_static.toFixed(2) + ' + ' + windForce.toFixed(2) + ' = ' + maxNormalForce.toFixed(2) + ' N'
        });
        step9.calculations.push({
            formula: conditionName + '拉伸应力 = |最大垂向载荷| / 胶层面积',
            value: 'σ_拉(' + conditionName + ') = |' + maxNormalForce.toFixed(2) + '| N / ' + results.bond_area.toFixed(2) + ' mm² = ' + tensileStress.toFixed(6) + ' MPa'
        });
        step9.calculations.push({
            formula: conditionName + '安全系数 = 老化后抗拉强度 / 拉伸应力',
            value: conditionName + '安全系数 = ' + results.aged_tensile_strength.toFixed(6) + ' MPa / ' + tensileStress.toFixed(6) + ' MPa = ' + safetyFactor.toFixed(2) + ' (要求: > ' + safetyFactorRequirements[conditionKey] + ')'
        });
    }
    steps.push(step9);

    // 安全判断
    results.is_shear_safe = results.shear_safety_factor > params.shear_req;
    results.is_tensile_safe = results.tensile_safety_factor > params.tensile_req;
    results.is_angle_safe = results.tan_theta <= params.allow_shear_angle;

    const allConditionsSafe = Object.values(results.condition_safety_status).every(Boolean);
    results.is_safe = results.is_shear_safe && results.is_tensile_safe && results.is_angle_safe && allConditionsSafe;

    results.steps = steps;

    return results;
}

function updateResultsDisplay(results, params) {
    const summaryDiv = document.getElementById('result-summary');
    summaryDiv.innerHTML = '';

    if (results.is_safe) {
        summaryDiv.className = 'result-summary status-safe';
        let statusHtml = '<div class="status-icon">✅</div>';
        statusHtml += '<h3>胶粘剂满足要求</h3>';
        statusHtml += '<p>所有安全指标均符合设计要求，可以使用该胶粘剂。</p>';
        statusHtml += '<p style="margin-top: 10px;">隧道及会车工况评估结果：</p>';

        for (const [condition, status] of Object.entries(results.condition_safety_status)) {
            statusHtml += '<p>' + (CONDITION_NAMES[condition] || condition) + ': ' + (status ? '满足 ✓' : '不满足 ✗') + '</p>';
        }
        summaryDiv.innerHTML = statusHtml;
    } else {
        summaryDiv.className = 'result-summary status-danger';
        const reasons = [];
        if (!results.is_shear_safe) reasons.push('剪切安全系数不足');
        if (!results.is_tensile_safe) reasons.push('拉伸安全系数不足');
        if (!results.is_angle_safe) reasons.push('剪切变形角超过允许值');

        for (const [condition, status] of Object.entries(results.condition_safety_status)) {
            if (!status) reasons.push((CONDITION_NAMES[condition] || condition) + '工况安全系数不足');
        }

        let statusHtml = '<div class="status-icon">❌</div>';
        statusHtml += '<h3>胶粘剂不满足要求</h3>';
        statusHtml += '<p>以下指标不满足要求：' + reasons.join('、') + '。</p>';
        statusHtml += '<p style="margin-top: 10px;">隧道及会车工况评估结果：</p>';

        for (const [condition, status] of Object.entries(results.condition_safety_status)) {
            statusHtml += '<p>' + (CONDITION_NAMES[condition] || condition) + ': ' + (status ? '满足 ✓' : '不满足 ✗') + '</p>';
        }
        summaryDiv.innerHTML = statusHtml;
    }

    // 更新三个结果卡片
    document.getElementById('shear-safety').textContent = results.shear_safety_factor.toFixed(2);
    document.getElementById('shear-safety').className = 'result-value ' + (results.is_shear_safe ? 'safe' : 'danger');
    document.getElementById('shear-requirement').textContent = '要求: > ' + params.shear_req;

    document.getElementById('tensile-safety').textContent = results.tensile_safety_factor.toFixed(2);
    document.getElementById('tensile-safety').className = 'result-value ' + (results.is_tensile_safe ? 'safe' : 'danger');
    document.getElementById('tensile-requirement').textContent = '要求: > ' + params.tensile_req;

    document.getElementById('shear-angle').textContent = results.tan_theta.toFixed(6);
    document.getElementById('shear-angle').className = 'result-value ' + (results.is_angle_safe ? 'safe' : 'danger');
    document.getElementById('angle-requirement').textContent = '要求: ≤ ' + params.allow_shear_angle;

    // 关键参数
    let keyParamsText = '胶层面积: ' + results.bond_area.toFixed(2) + ' mm²\n';
    keyParamsText += '剪切应力: ' + results.tau_shear.toFixed(6) + ' MPa\n';
    keyParamsText += '拉伸应力: ' + results.sigma_tensile.toFixed(6) + ' MPa\n';
    keyParamsText += '总老化系数: ' + results.total_aging_factor.toFixed(6) + '\n\n';
    keyParamsText += '老化后剪切强度: ' + results.aged_shear_strength.toFixed(6) + ' MPa\n';
    keyParamsText += '老化后抗拉强度: ' + results.aged_tensile_strength.toFixed(6) + ' MPa\n\n';
    keyParamsText += '总切向载荷: ' + results.f_shear_total.toFixed(2) + ' N\n';
    keyParamsText += '最大垂向载荷: ' + results.f_normal_max.toFixed(2) + ' N\n\n';
    keyParamsText += '工况安全系数:\n';

    for (const [condition, factor] of Object.entries(results.condition_safety_factors)) {
        const req = results.condition_requirements[condition];
        const status = results.condition_safety_status[condition];
        keyParamsText += (CONDITION_NAMES[condition] || condition) + ': ' + factor.toFixed(2) + ' (要求: > ' + req + ') ' + (status ? '✓' : '✗') + '\n';
    }

    document.getElementById('key-params').textContent = keyParamsText;

    // 详细计算过程
    const stepsDiv = document.getElementById('calculation-steps');
    stepsDiv.innerHTML = '';

    for (const step of results.steps) {
        const stepDiv = document.createElement('div');
        stepDiv.className = 'calculation-step';

        let stepHtml = '<h4>' + step.title + '</h4>';
        for (const calc of step.calculations) {
            stepHtml += '<div class="formula">● 公式: ' + calc.formula + '</div>';
            stepHtml += '<div class="value">● 计算: ' + calc.value + '</div>';
        }
        stepDiv.innerHTML = stepHtml;
        stepsDiv.appendChild(stepDiv);
    }
}

function resetParameters() {
    const defaults = {
        project_name: '',
        drawing_number: '',
        adhesive_model: '',
        glass_mass: 48,
        glass_area: 1,
        glass_length: 1084,
        glass_width: 967,
        installation_angle: 66,
        bond_width: 25,
        adhesive_thickness: 8,
        train_speed: 300,
        wind_pressure: 4500,
        long_acceleration: 3,
        trans_acceleration: 1,
        vert_acceleration: 3,
        min_temperature: -50,
        max_temperature: 45,
        glass_expansion: '9.28e-6',
        frame_expansion: '23e-6',
        frp_expansion: '15e-6',
        tensile_strength: 7.8,
        shear_strength: 5.3,
        allow_shear_angle: 0.5,
        temp_factor: 0.74,
        medium_factor: 0.91,
        geo_factor: 0.92,
        long_static_factor: 0.72,
        dynamic_factor: 0.83,
        human_factor: 0.95,
        shear_safety_req: 1.5,
        tensile_safety_req: 1.5,
        front_outdoor_req: 1.5,
        front_tunnel_entry_req: 1.5,
        front_tunnel_stable_req: 1.5,
        front_tunnel_exit_req: 2.0,
        front_overtaking_req: 2.0,
        side_outdoor_req: 2.0,
        side_tunnel_entry_req: 2.5,
        side_tunnel_stable_req: 2.0,
        side_tunnel_exit_req: 3.0,
        side_overtaking_req: 3.0
    };

    for (const [key, value] of Object.entries(defaults)) {
        const element = document.getElementById(key);
        if (element) {
            element.value = value;
        }
    }

    window.calculationResults = null;

    const summaryDiv = document.getElementById('result-summary');
    summaryDiv.className = 'result-summary status-pending';
    summaryDiv.innerHTML = '<div class="status-icon">⏳</div><h3>等待计算</h3><p>请输入参数并点击"开始计算"</p>';

    document.getElementById('shear-safety').textContent = '-';
    document.getElementById('shear-safety').className = 'result-value';
    document.getElementById('shear-requirement').textContent = '要求: -';

    document.getElementById('tensile-safety').textContent = '-';
    document.getElementById('tensile-safety').className = 'result-value';
    document.getElementById('tensile-requirement').textContent = '要求: -';

    document.getElementById('shear-angle').textContent = '-';
    document.getElementById('shear-angle').className = 'result-value';
    document.getElementById('angle-requirement').textContent = '要求: -';

    document.getElementById('key-params').textContent = '';
    document.getElementById('calculation-steps').innerHTML = '';
}

// Service Worker 注册（可选，用于PWA）
if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
        navigator.serviceWorker.register('service-worker.js')
            .then(registration => {
                console.log('ServiceWorker 注册成功:', registration);
            })
            .catch(err => {
                console.log('ServiceWorker 注册失败:', err);
            });
    });
}
