/**
 * Test function.
 * Does NOT fetch anything.
 */
function test() {
    return 12345;
}


/**
 * Test asynchronous function.
 * Does NOT fetch anything.
 */
async function asyncTest() {
    return 67890;
}


/**
 * NAV diagnostic.
 * Does NOT fetch anything yet.
 */
async function nav(schemeCode) {
    return 999;
}


/**
 * NAME diagnostic.
 */
async function name(schemeCode) {
    return "TEST NAME";
}


/**
 * ISIN diagnostic.
 */
async function isin(schemeCode) {
    return "TEST ISIN";
}


/**
 * PLAN diagnostic.
 */
async function plan(schemeCode) {
    return "TEST PLAN";
}


/**
 * OPTION diagnostic.
 */
async function option(schemeCode) {
    return "TEST OPTION";
}


/**
 * DATE diagnostic.
 */
async function navDate(schemeCode) {
    return "TEST DATE";
}


CustomFunctions.associate("NAV", nav);
CustomFunctions.associate("NAME", name);
CustomFunctions.associate("ISIN", isin);
CustomFunctions.associate("PLAN", plan);
CustomFunctions.associate("OPTION", option);
CustomFunctions.associate("DATE", navDate);
