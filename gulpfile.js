'use strict';

process.env.BROWSERSLIST_IGNORE_OLD_DATA = "true";

const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);
build.addSuppression(/Warning/gi);

build.eslint = build.eslint || {};
build.eslint.enabled = false;
if (build.tslintCmd) {
  build.tslintCmd.enabled = false;
}

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

build.initialize(require('gulp'));
