'use strict';

const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);



build.addSuppression(/Admins can make this solution available to all sites immediately, but the solution also contains feature\.xml elements for provisioning\./);
build.addSuppression(/Warning/gi);

build.initialize(require('gulp'));
