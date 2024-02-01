 import { createDarkTheme, createLightTheme } from '@fluentui/react-components';

 import type { BrandVariants, Theme } from '@fluentui/react-components';

 const companyBlue: BrandVariants = { 
   10: "#002436",
   20: "#0F1921",
   30: "#12293B",
   40: "#11354F",
   50: "#0D4264",
   60: "#03507A",
   70: "#225D86",
   80: "#3A6A90",
   90: "#4E779A",
   100: "#6184A5",
   110: "#7492AF",
   120: "#86A0BA",
   130: "#99AEC4",
   140: "#ABBDCF",
   150: "#BECBD9",
   160: "#D0DAE4"
};

export const lightThemeCompanyBlue: Theme = {
   ...createLightTheme(companyBlue), 
};

export const darkThemeCompanyBlue: Theme = {
   ...createDarkTheme(companyBlue), 
};

darkThemeCompanyBlue.colorBrandForeground1 = companyBlue[110]; // use brand[110] instead of brand[100]
darkThemeCompanyBlue.colorBrandForeground2 = companyBlue[120]; // use brand[120] instead of brand[110]
