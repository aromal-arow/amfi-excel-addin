function nav(schemeCode) {
    return 999;
}

function name(schemeCode) {
    return "TEST NAME";
}

function isin(schemeCode) {
    return "TEST ISIN";
}

function plan(schemeCode) {
    return "TEST PLAN";
}

function option(schemeCode) {
    return "TEST OPTION";
}

function navDate(schemeCode) {
    return "TEST DATE";
}


CustomFunctions.associate("NAV", nav);
CustomFunctions.associate("NAME", name);
CustomFunctions.associate("ISIN", isin);
CustomFunctions.associate("PLAN", plan);
CustomFunctions.associate("OPTION", option);
CustomFunctions.associate("DATE", navDate);
