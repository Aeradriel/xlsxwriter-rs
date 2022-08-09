use crate::{Format, FormatColor};

use super::Workbook;
use std::ffi::CString;

pub enum ConditionnalType {
    Cell,
    Text,
    TimePeriod,
    Average,
    Duplicate,
    Unique,
    Top,
    Bottom,
    Blanks,
    NoBlanks,
    Errors,
    NoErrors,
    Formula,
    TwoColorScale,
    ThreeColorScale,
    DataBar,
    IconSets,
}

pub enum ConditionnalCriteria {
    EqualTo,
    NotEqualTo,
    GreaterThan,
    LessThan,
    GreaterThanOrEqualTo,
    LessThanOrEqualTo,
    Between,
    NotBetween,
    TextContaining,
    TextNotContaining,
    TextBeginsWith,
    TextEndsWith,
    TimePeriodYesterday,
    TimePeriodToday,
    TimePeriodTomorrow,
    TimePeriodLastSevenDays,
    TimePeriodLastWeek,
    TimePeriodThisWeek,
    TimePeriodLastMonth,
    TimePeriodThisMonth,
    TimePeriodNextMonth,
    AverageAbove,
    AverageBelow,
    AverageAboveOrEqual,
    AverageBelowOrEqual,
    AverageOneStdDevAbove,
    AverageOneStdDevBelow,
    AverageTwoStdDevAbove,
    AverageTwoStdDevBelow,
    AverageThreeStdDevAbove,
    AverageThreeStdDevBelow,
    TopOrBottomPercent,
}

pub enum ConditionnalRuleType {
    Minimum,
    Number,
    Percent,
    Percentile,
    Formula,
    Maximum,
}

pub enum ConditionnalBarDirection {
    Context,
    RightToLeft,
    LeftToRight,
}

pub enum ConditionnalBarAxisPosition {
    Automatic,
    Midpoint,
    None,
}

pub enum ConditionnalIconType {
    ThreeArrowsColored,
    ThreeArrowsGray,
    ThreeFlags,
    ThreeTrafficLightsUnrimmed,
    ThreeTrafficLightsRimmed,
    ThreeSigns,
    ThreeSymbolsCircled,
    FourArrowsColored,
    FourArrowsGray,
    FourRedToBlack,
    FourRating,
    FourTrafficLights,
    FiveArrowsColored,
    FiveArrowsGray,
    FiveRatings,
    FiveQuarters,
}

pub struct ConditionnalFormat<'a> {
    pub _internal_format: *mut libxlsxwriter_sys::lxw_conditional_format,
    pub conditionnal_type: ConditionnalType,
    pub criteria: ConditionnalCriteria,
    pub value: f64,
    pub value_string: Option<String>,
    pub format: Format<'a>,
    pub min_value: f64,
    pub min_value_string: Option<String>,
    pub min_rule_type: ConditionnalRuleType,
    pub min_color: FormatColor,
    pub mid_value: f64,
    pub mid_value_string: Option<String>,
    pub mid_rule_type: ConditionnalRuleType,
    pub mid_color: FormatColor,
    pub max_value: f64,
    pub max_value_string: Option<String>,
    pub max_rule_type: ConditionnalRuleType,
    pub max_color: FormatColor,
    pub bar_color: FormatColor,
    pub bar_only: bool,
    pub data_bar_2010: bool,
    pub bar_solid: bool,
    pub bar_negative_color: FormatColor,
    pub bar_border_color: FormatColor,
    pub bar_negative_border_color: FormatColor,
    pub bar_negative_color_same: bool,
    pub bar_negative_border_color_same: bool,
    pub bar_no_border: bool,
    pub bar_direction: ConditionnalBarDirection,
    pub bar_axis_position: ConditionnalBarAxisPosition,
    pub bar_axis_color: FormatColor,
    pub icon_style: ConditionnalIconType,
    pub reverse_icons: bool,
    pub icons_only: bool,
    pub multi_range: Option<String>,
    pub stop_if_true: bool,
}

impl<'a> ConditionnalFormat<'a> {
    // TODO
}
