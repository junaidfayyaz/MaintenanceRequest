// src/utils/defaultHelper.ts
export var applyDefaults = function (properties, defaults) {
    Object.entries(defaults).forEach(function (_a) {
        var key = _a[0], value = _a[1];
        if (properties[key] === undefined ||
            properties[key] === null ||
            properties[key] === "") {
            properties[key] = value;
        }
    });
};
//# sourceMappingURL=defaultHelper.js.map