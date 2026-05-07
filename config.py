"""Configuration for offline conversion report automation.

Each OUTPUT entry defines one report file. To add a new vertical/output:
just add an entry — no code changes needed.
"""

# Time rules — applied to every output
SIGNUP_TIME = "23:55:00"
FTD_TIME = "23:58:00"
CURRENCY = "USD"

# Channel → report column "Type" values that belong to it
CHANNEL_TYPES = {
    "google": ["gclid", "gbraid"],
    "bing": ["msclkid"],
}

# Bing Cid prefix to strip
BING_CID_PREFIX = "m_"

# Output definitions. Each one produces one .xlsx file.
#
# Fields:
#   template:       template filename in templates/
#   output_prefix:  prefix for the generated file name
#   site_ids:       list of SiteIds belonging to this output
#   channel:        "google" or "bing"
#   signup_label:   exact wording for the signup row's Conversion Name
#                   (taken from each template to match what the AM expects)
#   ftd_label:      exact wording for the bare "ftds" row (non-casino)
#   casino_label:   exact wording for the collapsed casino row
#   brand_prefix:   prefix used when writing per-brand rows (e.g. "offline - ")
OUTPUTS = [
    {
        "key": "sport_google",
        "template": "Sport_Conversion_Google_-_Template.xlsx",
        "output_prefix": "Sport_Conversion_Google",
        "site_ids": [24],
        "channel": "google",
        "signup_label": "offline conversion - signup",
        "ftd_label": "offline - ftds",
        "casino_label": "offline - ftdcasino",
        "brand_prefix": "offline - ",
    },
    {
        "key": "sport_bing",
        "template": "Sport_Conversion_Bing_-_Template.xlsx",
        "output_prefix": "Sport_Conversion_Bing",
        "site_ids": [24],
        "channel": "bing",
        "signup_label": "offline conversion - signup",
        "ftd_label": "offline - ftds",
        "casino_label": "offline - ftdcasino",
        "brand_prefix": "offline - ",
    },
    {
        "key": "bingo_google",
        "template": "Bingo_Conversion_Google_-_Template.xlsx",
        "output_prefix": "Bingo_Conversion_Google",
        "site_ids": [568],
        "channel": "google",
        "signup_label": "offline conversion - signup",
        "ftd_label": "offline - ftds bingo",
        "casino_label": "offline - ftdscasino",
        "brand_prefix": "offline - ",
    },
    {
        "key": "bingo_bing",
        "template": "Bingo_Conversion_Bing_-_Template.xlsx",
        "output_prefix": "Bingo_Conversion_Bing",
        "site_ids": [568],
        "channel": "bing",
        "signup_label": "offline - signup",
        "ftd_label": "offline - ftds",
        "casino_label": "offline - ftdcasino",
        "brand_prefix": "offline - ",
    },
    {
        "key": "uk1_google",
        "template": "Offline_UK__1__Google_-_Template.xlsx",
        "output_prefix": "Offline_UK_1_Google",
        "site_ids": [25],
        "channel": "google",
        "signup_label": "offline conversion - signup",
        "ftd_label": "offline - ftds",
        # UK1/2/3 are casino-only verticals — but per AM clarification they still
        # use the 2-row treatment (per-brand row + ftds row), not the collapse.
        # casino_label is unused for these since brands are listed individually.
        "casino_label": "offline - ftdcasino",
        "brand_prefix": "offline - ",
    },
    {
        "key": "uk2_bing",
        "template": "UK_convftds__-_BING_-_UK__2_-_Template.xlsx",
        "output_prefix": "UK_convftds_BING_UK_2",
        "site_ids": [39],
        "channel": "bing",
        "signup_label": "offline - signup",
        "ftd_label": "offline - ftds",
        "casino_label": "offline - ftdcasino",
        "brand_prefix": "offline - ",
    },
    {
        "key": "uk3_bing",
        "template": "UK_convftds__-_BING_-_UK__3_-_Template.xlsx",
        "output_prefix": "UK_convftds_BING_UK_3",
        "site_ids": [25],
        "channel": "bing",
        "signup_label": "offline - signup",
        "ftd_label": "offline - ftds",
        "casino_label": "offline - ftdcasino",
        "brand_prefix": "offline - ",
    },
]

# Outputs where casino brands collapse to one row (sport, bingo).
# UK1/UK2/UK3 are casino-only but use 2-row treatment per AM, so they're NOT here.
COLLAPSE_CASINO_OUTPUTS = {"sport_google", "sport_bing", "bingo_google", "bingo_bing"}

# Track360 raw report column names
COL_SITE_ID = "TrackingParams - SiteId"
COL_TYPE = "TrackingParams - Type"
COL_BRAND = "TrackingParams - BrandName"
COL_CID = "TrackingParams - Cid"
COL_SIGNUP_DATE = "Signup Date"
COL_FTD_DATE = "FtdDate"
COL_SIGNUPS = "Signups"
COL_FTDS = "FTDs"
COL_REVENUES = "Revenues"

REQUIRED_COLUMNS = [
    COL_SITE_ID, COL_TYPE, COL_BRAND, COL_CID,
    COL_SIGNUP_DATE, COL_FTD_DATE,
    COL_SIGNUPS, COL_FTDS, COL_REVENUES,
]
