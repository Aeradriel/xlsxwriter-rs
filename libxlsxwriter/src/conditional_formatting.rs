use std::{
    ffi::{c_char, CStr, CString},
    ptr::null_mut,
};

use crate::{convert_bool, Format, FormatColor};

#[derive(Debug)]
pub enum ConditionalType {
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

#[derive(Debug)]
pub enum ConditionalCriteria {
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

#[derive(Debug)]
pub enum ConditionalRuleType {
    Minimum,
    Number,
    Percent,
    Percentile,
    Formula,
    Maximum,
}

#[derive(Debug)]
pub enum ConditionalBarDirection {
    Context,
    RightToLeft,
    LeftToRight,
}

#[derive(Debug)]
pub enum ConditionalBarAxisPosition {
    Automatic,
    Midpoint,
    None,
}

#[derive(Debug)]
pub enum ConditionalIconType {
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

#[derive(Debug)]
pub struct ConditionalFormat {
    pub _internal_format: libxlsxwriter_sys::lxw_conditional_format,
    string_value: Option<Vec<u8>>,
}

impl ConditionalType {
    pub fn value(&self) -> u8 {
        let value = match self {
            ConditionalType::Cell => {
                libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_TYPE_CELL
            }
            ConditionalType::Text => {
                libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_TYPE_TEXT
            }
            ConditionalType::TimePeriod => {
                libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_TYPE_TIME_PERIOD
            }
            ConditionalType::Average => {
                libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_TYPE_AVERAGE
            }
            ConditionalType::Duplicate => {
                libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_TYPE_DUPLICATE
            }
            ConditionalType::Unique => {
                libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_TYPE_UNIQUE
            }
            ConditionalType::Top => {
                libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_TYPE_TOP
            }
            ConditionalType::Bottom => {
                libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_TYPE_BOTTOM
            }
            ConditionalType::Blanks => {
                libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_TYPE_BLANKS
            }
            ConditionalType::NoBlanks => {
                libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_TYPE_NO_BLANKS
            }
            ConditionalType::Errors => {
                libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_TYPE_ERRORS
            }
            ConditionalType::NoErrors => {
                libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_TYPE_NO_ERRORS
            }
            ConditionalType::Formula => {
                libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_TYPE_FORMULA
            }
            ConditionalType::TwoColorScale => {
                libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_2_COLOR_SCALE
            }
            ConditionalType::ThreeColorScale => {
                libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_3_COLOR_SCALE
            }
            ConditionalType::DataBar => {
                libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_DATA_BAR
            }
            ConditionalType::IconSets => {
                libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_TYPE_ICON_SETS
            }
        };
        value as u8
    }
}

impl ConditionalCriteria {
    pub fn value(&self) -> u8 {
        let value = match self {
            ConditionalCriteria::EqualTo => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_EQUAL_TO,
            ConditionalCriteria::NotEqualTo => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_NOT_EQUAL_TO,
            ConditionalCriteria::GreaterThan => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_GREATER_THAN,
            ConditionalCriteria::LessThan => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_LESS_THAN,
            ConditionalCriteria::GreaterThanOrEqualTo => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_GREATER_THAN_OR_EQUAL_TO,
            ConditionalCriteria::LessThanOrEqualTo => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_LESS_THAN_OR_EQUAL_TO,
            ConditionalCriteria::Between => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_BETWEEN,
            ConditionalCriteria::NotBetween => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_NOT_BETWEEN,
            ConditionalCriteria::TextContaining => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TEXT_CONTAINING,
            ConditionalCriteria::TextNotContaining => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TEXT_NOT_CONTAINING,
            ConditionalCriteria::TextBeginsWith => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TEXT_BEGINS_WITH,
            ConditionalCriteria::TextEndsWith => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TEXT_ENDS_WITH,
            ConditionalCriteria::TimePeriodYesterday => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_YESTERDAY,
            ConditionalCriteria::TimePeriodToday => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_TODAY,
            ConditionalCriteria::TimePeriodTomorrow => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_TOMORROW,
            ConditionalCriteria::TimePeriodLastSevenDays => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_7_DAYS,
            ConditionalCriteria::TimePeriodLastWeek => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_WEEK,
            ConditionalCriteria::TimePeriodThisWeek => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_THIS_WEEK,
            ConditionalCriteria::TimePeriodLastMonth => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_MONTH,
            ConditionalCriteria::TimePeriodThisMonth => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_THIS_MONTH,
            ConditionalCriteria::TimePeriodNextMonth => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_MONTH,
            ConditionalCriteria::AverageAbove => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_AVERAGE_ABOVE,
            ConditionalCriteria::AverageBelow => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_AVERAGE_BELOW,
            ConditionalCriteria::AverageAboveOrEqual => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_AVERAGE_ABOVE_OR_EQUAL,
            ConditionalCriteria::AverageBelowOrEqual => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_AVERAGE_BELOW_OR_EQUAL,
            ConditionalCriteria::AverageOneStdDevAbove => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_AVERAGE_1_STD_DEV_ABOVE,
            ConditionalCriteria::AverageOneStdDevBelow => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_AVERAGE_1_STD_DEV_BELOW,
            ConditionalCriteria::AverageTwoStdDevAbove => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_AVERAGE_2_STD_DEV_ABOVE,
            ConditionalCriteria::AverageTwoStdDevBelow => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_AVERAGE_2_STD_DEV_BELOW,
            ConditionalCriteria::AverageThreeStdDevAbove => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_ABOVE,
            ConditionalCriteria::AverageThreeStdDevBelow => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_BELOW,
            ConditionalCriteria::TopOrBottomPercent => libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TOP_OR_BOTTOM_PERCENT,
        };
        value as u8
    }
}

impl ConditionalRuleType {
    pub fn value(&self) -> u8 {
        let value = match self {
            ConditionalRuleType::Minimum => libxlsxwriter_sys::lxw_conditional_format_rule_types_LXW_CONDITIONAL_RULE_TYPE_MINIMUM,
            ConditionalRuleType::Number => libxlsxwriter_sys::lxw_conditional_format_rule_types_LXW_CONDITIONAL_RULE_TYPE_NUMBER,
            ConditionalRuleType::Percent => libxlsxwriter_sys::lxw_conditional_format_rule_types_LXW_CONDITIONAL_RULE_TYPE_PERCENT,
            ConditionalRuleType::Percentile => libxlsxwriter_sys::lxw_conditional_format_rule_types_LXW_CONDITIONAL_RULE_TYPE_PERCENTILE,
            ConditionalRuleType::Formula => libxlsxwriter_sys::lxw_conditional_format_rule_types_LXW_CONDITIONAL_RULE_TYPE_FORMULA,
            ConditionalRuleType::Maximum => libxlsxwriter_sys::lxw_conditional_format_rule_types_LXW_CONDITIONAL_RULE_TYPE_MAXIMUM
        };
        value as u8
    }
}

impl ConditionalBarDirection {
    pub fn value(&self) -> u8 {
        let value = match self {
            ConditionalBarDirection::Context => libxlsxwriter_sys::lxw_conditional_format_bar_direction_LXW_CONDITIONAL_BAR_DIRECTION_CONTEXT,
            ConditionalBarDirection::RightToLeft => libxlsxwriter_sys::lxw_conditional_format_bar_direction_LXW_CONDITIONAL_BAR_DIRECTION_RIGHT_TO_LEFT,
            ConditionalBarDirection::LeftToRight => libxlsxwriter_sys::lxw_conditional_format_bar_direction_LXW_CONDITIONAL_BAR_DIRECTION_LEFT_TO_RIGHT,
        };
        value as u8
    }
}

impl ConditionalBarAxisPosition {
    pub fn value(&self) -> u8 {
        let value = match self {
            ConditionalBarAxisPosition::Automatic => libxlsxwriter_sys::lxw_conditional_bar_axis_position_LXW_CONDITIONAL_BAR_AXIS_AUTOMATIC,
            ConditionalBarAxisPosition::Midpoint => libxlsxwriter_sys::lxw_conditional_bar_axis_position_LXW_CONDITIONAL_BAR_AXIS_MIDPOINT,
            ConditionalBarAxisPosition::None => libxlsxwriter_sys::lxw_conditional_bar_axis_position_LXW_CONDITIONAL_BAR_AXIS_NONE,
        };
        value as u8
    }
}

impl ConditionalIconType {
    pub fn value(&self) -> u8 {
        let value = match self {
            ConditionalIconType::ThreeArrowsColored => libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_3_ARROWS_COLORED,
            ConditionalIconType::ThreeArrowsGray => libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_3_ARROWS_GRAY,
            ConditionalIconType::ThreeFlags => libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_3_FLAGS,
            ConditionalIconType::ThreeTrafficLightsUnrimmed => libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_3_TRAFFIC_LIGHTS_UNRIMMED,
            ConditionalIconType::ThreeTrafficLightsRimmed => libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_3_TRAFFIC_LIGHTS_RIMMED,
            ConditionalIconType::ThreeSigns => libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_3_SIGNS,
            ConditionalIconType::ThreeSymbolsCircled => libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_3_SYMBOLS_CIRCLED,
            ConditionalIconType::FourArrowsColored => libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_4_ARROWS_COLORED,
            ConditionalIconType::FourArrowsGray => libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_4_ARROWS_GRAY,
            ConditionalIconType::FourRedToBlack => libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_4_RED_TO_BLACK,
            ConditionalIconType::FourRating => libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_4_RATINGS,
            ConditionalIconType::FourTrafficLights => libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_4_TRAFFIC_LIGHTS,
            ConditionalIconType::FiveArrowsColored => libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_5_ARROWS_COLORED,
            ConditionalIconType::FiveArrowsGray => libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_5_ARROWS_GRAY,
            ConditionalIconType::FiveRatings => libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_5_RATINGS,
            ConditionalIconType::FiveQuarters => libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_5_QUARTERS,
        };
        value as u8
    }
}

impl ConditionalFormat {
    pub fn new(format: Format) -> Self {
        let internal_format = libxlsxwriter_sys::lxw_conditional_format {
            type_: libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_TYPE_CELL as u8,
            criteria: libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_EQUAL_TO
                as u8,
            value: 0.0,
            value_string: null_mut(),
            format: format.format,
            min_value: 0.0,
            min_value_string: null_mut(),
            min_rule_type: libxlsxwriter_sys::lxw_conditional_format_rule_types_LXW_CONDITIONAL_RULE_TYPE_NUMBER as u8,
            min_color: libxlsxwriter_sys::lxw_defined_colors_LXW_COLOR_BLACK,
            mid_value: 0.0,
            mid_value_string: null_mut(),
            mid_rule_type: libxlsxwriter_sys::lxw_conditional_format_rule_types_LXW_CONDITIONAL_RULE_TYPE_NUMBER as u8,
            mid_color: libxlsxwriter_sys::lxw_defined_colors_LXW_COLOR_BLACK,
            max_value: 0.0,
            max_value_string: null_mut(),
            max_rule_type: libxlsxwriter_sys::lxw_conditional_format_rule_types_LXW_CONDITIONAL_RULE_TYPE_NUMBER as u8,
            max_color: libxlsxwriter_sys::lxw_defined_colors_LXW_COLOR_BLACK,
            bar_color: libxlsxwriter_sys::lxw_defined_colors_LXW_COLOR_BLACK,
            bar_only: 0,
            data_bar_2010: 0,
            bar_solid: 0,
            bar_negative_color: libxlsxwriter_sys::lxw_defined_colors_LXW_COLOR_BLACK,
            bar_border_color: libxlsxwriter_sys::lxw_defined_colors_LXW_COLOR_BLACK,
            bar_negative_border_color: libxlsxwriter_sys::lxw_defined_colors_LXW_COLOR_BLACK,
            bar_negative_color_same: 0,
            bar_negative_border_color_same: 0,
            bar_no_border: 0,
            bar_direction: libxlsxwriter_sys::lxw_conditional_format_bar_direction_LXW_CONDITIONAL_BAR_DIRECTION_CONTEXT as u8,
            bar_axis_position: libxlsxwriter_sys::lxw_conditional_bar_axis_position_LXW_CONDITIONAL_BAR_AXIS_AUTOMATIC as u8,
            bar_axis_color: libxlsxwriter_sys::lxw_defined_colors_LXW_COLOR_BLACK,
            icon_style: libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_3_ARROWS_COLORED as u8,
            reverse_icons: 0,
            icons_only: 0,
            multi_range: null_mut(),
            stop_if_true: 0,
        };

        ConditionalFormat {
            _internal_format: internal_format,
            string_value: None,
        }
    }

    pub fn set_conditional_type(mut self, conditional_type: ConditionalType) -> Self {
        self._internal_format.type_ = conditional_type.value() as u8;
        self
    }

    pub fn set_criteria(mut self, criteria: ConditionalCriteria) -> Self {
        self._internal_format.criteria = criteria.value() as u8;
        self
    }

    pub fn set_value(mut self, value: f64) -> Self {
        self._internal_format.value = value;
        self
    }

    pub fn set_value_string(mut self, value_string: Option<String>) -> Self {
        self.string_value = option_str_to_cstr_bytes(&value_string);
        self._internal_format.value_string =
            self.string_value
                .as_mut()
                .map(|x| x.as_mut_ptr())
                .unwrap_or(std::ptr::null_mut()) as *mut c_char;
        self
    }

    pub fn set_format(mut self, format: &Format) -> Self {
        self._internal_format.format = format.format;
        self
    }

    pub fn set_min_value(mut self, min_value: f64) -> Self {
        self._internal_format.min_value = min_value;
        self
    }

    pub fn set_min_value_string(mut self, min_value_string: Option<String>) -> Self {
        self._internal_format.min_value_string = option_str_to_cstr_bytes(&min_value_string)
            .as_mut()
            .map(|x| x.as_mut_ptr())
            .unwrap_or(std::ptr::null_mut())
            as *mut c_char;
        self
    }

    pub fn set_min_rule_type(mut self, min_rule_type: ConditionalRuleType) -> Self {
        self._internal_format.min_rule_type = min_rule_type.value();
        self
    }

    pub fn set_min_color(mut self, min_color: FormatColor) -> Self {
        self._internal_format.min_color = min_color.value();
        self
    }

    pub fn set_mid_value(mut self, mid_value: f64) -> Self {
        self._internal_format.mid_value = mid_value;
        self
    }

    pub fn set_mid_value_string(mut self, mid_value_string: Option<String>) -> Self {
        self._internal_format.mid_value_string = option_str_to_cstr_bytes(&mid_value_string)
            .as_mut()
            .map(|x| x.as_mut_ptr())
            .unwrap_or(std::ptr::null_mut())
            as *mut c_char;
        self
    }

    pub fn set_mid_rule_type(mut self, mid_rule_type: ConditionalRuleType) -> Self {
        self._internal_format.mid_rule_type = mid_rule_type.value();
        self
    }

    pub fn set_mid_color(mut self, mid_color: FormatColor) -> Self {
        self._internal_format.mid_color = mid_color.value();
        self
    }

    pub fn set_max_value(mut self, max_value: f64) -> Self {
        self._internal_format.max_value = max_value;
        self
    }

    pub fn set_max_value_string(mut self, max_value_string: Option<String>) -> Self {
        self._internal_format.max_value_string = option_str_to_cstr_bytes(&max_value_string)
            .as_mut()
            .map(|x| x.as_mut_ptr())
            .unwrap_or(std::ptr::null_mut())
            as *mut c_char;
        self
    }

    pub fn set_max_rule_type(mut self, max_rule_type: ConditionalRuleType) -> Self {
        self._internal_format.max_rule_type = max_rule_type.value();
        self
    }

    pub fn set_max_color(mut self, max_color: FormatColor) -> Self {
        self._internal_format.max_color = max_color.value();
        self
    }

    pub fn set_bar_color(mut self, bar_color: FormatColor) -> Self {
        self._internal_format.bar_color = bar_color.value();
        self
    }

    pub fn set_bar_only(mut self, bar_only: bool) -> Self {
        self._internal_format.bar_only = convert_bool(bar_only);
        self
    }

    pub fn set_data_bar_2010(mut self, data_bar_2010: bool) -> Self {
        self._internal_format.data_bar_2010 = convert_bool(data_bar_2010);
        self
    }

    pub fn set_bar_solid(mut self, bar_solid: bool) -> Self {
        self._internal_format.bar_solid = convert_bool(bar_solid);
        self
    }

    pub fn set_bar_negative_color(mut self, bar_negative_color: FormatColor) -> Self {
        self._internal_format.bar_negative_color = bar_negative_color.value();
        self
    }

    pub fn set_bar_border_color(mut self, bar_border_color: FormatColor) -> Self {
        self._internal_format.bar_border_color = bar_border_color.value();
        self
    }

    pub fn set_bar_negative_border_color(mut self, bar_negative_border_color: FormatColor) -> Self {
        self._internal_format.bar_negative_border_color = bar_negative_border_color.value();
        self
    }

    pub fn set_bar_negative_color_same(mut self, bar_negative_color_same: bool) -> Self {
        self._internal_format.bar_negative_color_same = convert_bool(bar_negative_color_same);
        self
    }

    pub fn set_bar_negative_border_color_same(
        mut self,
        bar_negative_border_color_same: bool,
    ) -> Self {
        self._internal_format.bar_negative_border_color_same =
            convert_bool(bar_negative_border_color_same);
        self
    }

    pub fn set_bar_no_border(mut self, bar_no_border: bool) -> Self {
        self._internal_format.bar_no_border = convert_bool(bar_no_border);
        self
    }

    pub fn set_bar_direction(mut self, bar_direction: ConditionalBarDirection) -> Self {
        self._internal_format.bar_direction = bar_direction.value();
        self
    }

    pub fn set_bar_axis_position(mut self, bar_axis_position: ConditionalBarAxisPosition) -> Self {
        self._internal_format.bar_axis_position = bar_axis_position.value();
        self
    }

    pub fn set_bar_axis_color(mut self, bar_axis_color: FormatColor) -> Self {
        self._internal_format.bar_axis_color = bar_axis_color.value();
        self
    }

    pub fn set_icon_style(mut self, icon_style: ConditionalIconType) -> Self {
        self._internal_format.icon_style = icon_style.value();
        self
    }

    pub fn set_reverse_icons(mut self, reverse_icons: bool) -> Self {
        self._internal_format.reverse_icons = convert_bool(reverse_icons);
        self
    }

    pub fn set_icons_only(mut self, icons_only: bool) -> Self {
        self._internal_format.icons_only = convert_bool(icons_only);
        self
    }

    pub fn set_multi_range(mut self, multi_range: Option<String>) -> Self {
        self._internal_format.multi_range = option_str_to_cstr_bytes(&multi_range)
            .as_mut()
            .map(|x| x.as_mut_ptr())
            .unwrap_or(std::ptr::null_mut())
            as *mut c_char;
        self
    }

    pub fn set_stop_if_true(mut self, stop_if_true: bool) -> Self {
        self._internal_format.stop_if_true = convert_bool(stop_if_true);
        self
    }
}

fn option_str_to_cstr_bytes(s: &Option<String>) -> Option<Vec<u8>> {
    s.as_ref().map(|x| {
        CString::new(x as &str)
            .unwrap()
            .into_bytes_with_nul()
            .to_vec()
    })
}
