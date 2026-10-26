
function buildSoftSet(data, rules, objectColumn) {
    const softSet = {};

    for (const [parameter, rule] of Object.entries(rules)) {
        softSet[parameter] = data
            .filter(row => {
                const cell = row[parameter];

                if (rule.type === "numeric") {
                    const value = Number(cell);
                    const target = Number(rule.value);

                    if (cell === "" || cell == null || !Number.isFinite(value)) {
                        return false;
                    }

                    switch (rule.operator) {
                        case "≥": return value >= target;
                        case "≤": return value <= target;
                        case ">": return value > target;
                        case "<": return value < target;
                        case "==": return value === target;
                        default: return false;
                    }
                }

                return (rule.values || []).includes(String(cell));
            })
            .map(row => String(row[objectColumn] ?? ""))
            .filter(Boolean);
    }

    return softSet;
}

function recommendSoftSet(data, objectColumn, rules, preferences) {
    if (!data?.length || !objectColumn) {
        throw new Error("Dataset belum tersedia.");
    }

    if (!preferences?.length) {
        throw new Error("Pilih minimal satu parameter preferensi.");
    }

    const softSet = buildSoftSet(data, rules, objectColumn);

    const objects = data
        .map(row => row[objectColumn])
        .filter(value => value != null && String(value).trim() !== "")
        .map(String);

    const result = objects.map(objectName => {
        const matched = preferences.filter(parameter =>
            (softSet[parameter] || []).includes(objectName)
        );

        return {
            Nama: objectName,
            Score: matched.length / preferences.length,
            "Parameter Terpenuhi": matched
        };
    });

    result.sort((a, b) => b.Score - a.Score);

    result.forEach((item, index) => {
        item.Peringkat = index + 1;
    });

    const maxScore = Math.max(0, ...result.map(item => item.Score));

    return {
        soft_set: softSet,
        result,
        best: result.filter(item => item.Score === maxScore),
        max_score: maxScore
    };
}
