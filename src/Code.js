// Entry point for the Apps Script bundle.
//
// The deployed code is a single file (dist/Code.gs), while the
// source is split into the ES modules below. build.js inlines
// them into one scope in exactly this order, so this list is
// the bundle's table of contents: a new module joins the
// bundle by adding one import line here. The tests import the
// modules directly and never go through this file.
import './common.js';
import './auth.js';
import './admin.js';
import './subscriptions.js';
import './article.js';
import './items.js';
import './room.js';
import './main.js';
