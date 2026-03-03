window.ADDRESS_MAPPINGS = {
    "states": {
        "Alabama": "AL",
        "Alaska": "AK",
        "American Samoa": "AS",
        "Arizona": "AZ",
        "Arkansas": "AR",
        "California": "CA",
        "Colorado": "CO",
        "Connecticut": "CT",
        "Delaware": "DE",
        "District Of Columbia": "DC",
        "Federated States Of Micronesia": "FM",
        "Florida": "FL",
        "Georgia": "GA",
        "Guam": "GU",
        "Hawaii": "HI",
        "Idaho": "ID",
        "Illinois": "IL",
        "Indiana": "IN",
        "Iowa": "IA",
        "Kansas": "KS",
        "Kentucky": "KY",
        "Louisiana": "LA",
        "Maine": "ME",
        "Marshall Islands": "MH",
        "Maryland": "MD",
        "Massachusetts": "MA",
        "Michigan": "MI",
        "Minnesota": "MN",
        "Mississippi": "MS",
        "Missouri": "MO",
        "Montana": "MT",
        "Nebraska": "NE",
        "Nevada": "NV",
        "New Hampshire": "NH",
        "New Jersey": "NJ",
        "New Mexico": "NM",
        "New York": "NY",
        "North Carolina": "NC",
        "North Dakota": "ND",
        "Northern Mariana Islands": "MP",
        "Ohio": "OH",
        "Oklahoma": "OK",
        "Oregon": "OR",
        "Palau": "PW",
        "Pennsylvania": "PA",
        "Puerto Rico": "PR",
        "Rhode Island": "RI",
        "South Carolina": "SC",
        "South Dakota": "SD",
        "Tennessee": "TN",
        "Texas": "TX",
        "Utah": "UT",
        "Vermont": "VT",
        "Virgin Islands": "VI",
        "Virginia": "VA",
        "Washington": "WA",
        "West Virginia": "WV",
        "Wisconsin": "WI",
        "Wyoming": "WY"
    },
    "countries": {
        "United States": "USA",
        "United States of America": "USA",
        "Canada": "CAN",
        "Mexico": "MEX",
        "United Kingdom": "GBR",
        "Australia": "AUS"
    }
};

window.getAbbreviation = function(type, value) {
    if (!value) return "";
    const map = window.ADDRESS_MAPPINGS[type];
    if (!map) return value;
    
    // Check direct match
    if (map[value]) return map[value];
    
    // Check case-insensitive
    const lowerValue = value.toLowerCase();
    for (const key in map) {
        if (key.toLowerCase() === lowerValue) return map[key];
    }
    
    // Return original if no match found (it might already be abbreviated)
    return value;
};
