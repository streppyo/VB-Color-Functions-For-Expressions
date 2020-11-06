# VB-Color-Functions-For-Expressions
Lightweight function scripts to put into SQL Server Reporting Services and other VB applications. All of the function in this script will return only Hex Color Codes as String, e.g. `#F0ADFF`. The color function can then be easily feed into expressions.

## Function List
### ColorMix
`ColorMix ( Value, MaxValue, MinValue, MaxValueColor [, MinValueColor] )`
A function that will interpolate intermediate colours between 2 specified colours according to _value_ provided. Can be used to create gradient colour scales for heatmap. 

Parameters | Type | Optional | Description
---|---|---|---
Value | Double | No | The value to be evaluated over _MaxValue_ and _MinValue_, usually between _MaxValue_ and _MinValue_. Values that overshoot the range will be truncated.
MaxValue | Double | No | The value that correspond to _MaxValueColor_. _Value_ larger than _MaxValue_ will gives color that specified in _MaxValueColor_. The closer the _value_ gets to the _MaxValue_ supplied, the closer the color to _MaxValueColor_.
MinValue | Double | No | The value that correspond to _MinValueColor_ or *white*. Works pretty the same (or opposite) to _MaxValue_, and should be smaller than the _MaxValue_.
MaxValueColor  | String | No | The Hex Colour Code for the colour that correspond to _MaxValue_. Example `#20ADFF`.
MinValueColor  | String | Yes |  The Hex Colour Code for the colour that correspond to _MinValue_. If no value is supplied, the _MinValue_ will correspond to white `#FFFFFF`).

### ColorMix2
`ColorMix2 ( Value, MaxValue, MinValue, MaxValueColor , MinValueColor, MidValueColor )`
A function that will interpolate intermediate colours between 3 specified colours according to _value_ provided. Can be used to create gradient colour scales with 2 sides for heatmaps or other plots (e.g. blue-white-red). It is just a two-sided version of the `ColorMix()` function. 

Parameters | Type | Optional | Description
---|---|---|---
Value | Double | No | The value to be evaluated over _MaxValue_ and _MinValue_. Values that overshoot the range will be truncated. If _value_ falls between _MaxValue_ and the middle of _MaxValue_ and _MinValue_, a `ColorMix` between Max and Mid Colour (default white) will be returned.
MaxValue | Double | No | The value that correspond to _MaxValueColor_. _Value_ larger than _MaxValue_ will gives color that specified in _MaxValueColor_. The closer the _value_ gets to the _MaxValue_ supplied, the closer the color to _MaxValueColor_.
MinValue | Double | No | The value that correspond to _MinValueColor_ or *white*. Works pretty the same (or opposite) to _MaxValue_.
MaxValueColor  | String | No | The Hex Colour Code for the colour that correspond to _MaxValue_. Example `#20ADFF`.
MinValueColor  | String | No |  The Hex Colour Code for the colour that correspond to _MinValue_. Example `#AAD3FF`.
MidValueColor  | String | Yes |  The Hex Colour Code for the colour that correspond to _MidValue_. If no value is supplied, the _MidValue_ will correspond to white `#FFFFFF`).

### RBG
`RGB ( red, blue, green)`
Converts a RGB color code into Hex Colour Code that can be used by applications as expressions.
Parameters | Type | Optional | Description
---|---|---|---
red | Double | No | Between 0 and 255. Non integral value will be rounded off and values that overshoot will be truncated. Represents the numerical value for red.
green | Double | No | Same as above (but green).
blue | Double | No | Same as above (but blue).

### HSL
`HSL ( hue, saturation, lightness)`
Converts a HSL color code into Hex Colour Code that can be used by applications as expressions.
Parameters | Type | Optional | Description
---|---|---|---
hue | Double | No | **Between 0 and 1**. Represents the degree of hue. Expecting to be larger than or equal to 0 and smailler than 1. 
saturation | Double | No | **Between 0 and 1**. Saturation of the colour. Values taht overshoot will be truncated.
lightness | Double | No | **Between 0 and 1**. Lightness of the colour. Values that overshoot will be truncated.

### HSV
`HSV ( hue, saturation, value)`
Converts a HSL color code into Hex Colour Code that can be used by applications as expressions.
Parameters | Type | Optional | Description
---|---|---|---
hue | Double | No | **Between 0 and 1**. Represents the degree of hue. Expecting to be larger than or equal to 0 and smailler than 1. 
saturation | Double | No | **Between 0 and 1**. Saturation of the colour. Values taht overshoot will be truncated.
value | Double | No | **Between 0 and 1**. Value (Bightness) of the colour. Values that overshoot will be truncated.

### HSL360, HSV360
`HSL360 ( hue, saturation, lightness)`
`HSV360 ( hue, saturation, value)`
Allows `HSL` and `HSV` to be called with values that ranged in `[0,360)` , `[0,100]` and  `[0,100]` respectively that will make some users to feed parameters more coveniently.

### RandomColor, 
`RandomColor ()`
Return a random colour that is evenly distributed over the RGB color space.
`RandomColorHSL ()`
Return a random colour that is evenly distributed over the HLS color space.
`RandomColorHSV ()`
Return a random colour that is evenly distributed over the HSV color space.

