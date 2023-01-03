import { extendTheme } from 'native-base'

export const THEME = extendTheme({
  colors: {
    primary: {
      700: '#996DFF',
    },
    secondary: {
      700: '#FBA94C',
    },
    green: {
      700: '#00875F',
      500: '#00B37E',
      300: '#04D361',
    },
    gray: {
      700: '#121214',
      600: '#202024',
      500: '#29292E',
      400: '#323238',
      300: '#7C7C8A',
      200: '#C4C4CC',
      100: '#E1E1E6',
    },
    white: '#FFFFFF',
  },
  fonts: {
    heading: 'Roboto_700Bold',
    body: 'Roboto_400Regular',
  },
  fontSizes: {
    xs: 12,
    sm: 14,
    md: 16,
    lg: 18,
    xl: 20,
  },
  sizes: {
    14: 56,
  },
  fontWeights: {
    hairline: 100,
    thin: 200,
    light: 300,
    normal: 400,
    medium: 500,
    semibold: 600,
    bold: 700,
    extrabold: 800,
    black: 900,
    extraBlack: 950,
  },
  letterSpacings: {
    xs: '-0.05em',
    sm: '-0.025em',
    md: 0,
    lg: '0.025em',
    xl: '0.3em',
  },
})
