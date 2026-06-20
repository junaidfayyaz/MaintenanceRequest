// src/utils/defaultHelper.ts
export const applyDefaults = (
  properties: { [key: string]: any },
  defaults: { [key: string]: any }
): void => {
  Object.entries(defaults).forEach(([key, value]) => {
    if (
      properties[key] === undefined ||
      properties[key] === null ||
      properties[key] === ""
    ) {
      properties[key] = value;
    }
  });
};
