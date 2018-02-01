# wait-until [![Build status](https://img.shields.io/travis/nylen/wait-until.svg?style=flat)](https://travis-ci.org/nylen/wait-until) [![npm package](http://img.shields.io/npm/v/wait-until.svg?style=flat)](https://www.npmjs.org/package/wait-until)

This simple Node.js module provides a way to check for a condition every so
often, then run a callback after the condition is met or the allotted time
expires.

This is good if you need to wait for specific state changes in code that you
don't have direct control over, like a headless web browser.

## Usage

```js
var waitUntil = require('wait-until');

waitUntil(interval, times, function condition() {
    return (someCondition ? true : false);
}, function done(result) {
    // result is true on success or false if the condition was never met
});
```

The `condition` function will be called up to `times` times, starting after
`interval` milliseconds.  Once it returns `true` (or a truthy value), the
`done` callback will be called with `result` equal to `true` (or whatever value
the `condition` function returned).

If the condition is never met within the specified timeframe, `done` will be
called with `result` equal to `false` (or whatever falsy value the `condition`
function returned the last time it was called).

### Fluent interface

**Don't want to remember the order of arguments?**  Then use the
[fluent interface](https://en.wikipedia.org/wiki/Fluent_interface):

```js
var waitUntil = require('wait-until');

waitUntil()
    .interval(500)
    .times(10)
    .condition(function() {
        return (someCondition ? true : false);
    })
    .done(function(result) {
        // do stuff
    });
```

### Async conditions

If the `condition` function accepts an argument, then it is assumed to be a
callback function which will be called with the result of the test:

```js
var waitUntil = require('wait-until');

waitUntil()
    .interval(500)
    .times(10)
    .condition(function(cb) {
        process.nextTick(function() {
            cb(someCondition ? true : false);
        });
    })
    .done(function(result) {
        // do stuff
    });
```

### Other notes

You can specify `.times(Infinity)` if you like, but it's probably better to set
and enforce a reasonable timeout.
