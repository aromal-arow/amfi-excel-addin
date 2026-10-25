/**
 * AMFI.SCHEMENAME
 */
async function schemeName(schemeCode) {

    try {

        const fund = await getFund(schemeCode);

        if (!fund) {
            return "Scheme not found";
        }

        return fund.name;

    } catch (error) {

        return "Unable to load AMFI data";
    }
}


CustomFunctions.associate("NAV", nav);
CustomFunctions.associate("SCHEMENAME", schemeName);
