import { fileURLToPath } from 'url';
import path from 'path';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const cfg = {
  USERNAME: 'sanjay.shetty@dextero.in',
  PASSWORD: 'BNI@2429',

  LOGIN_URL:  'https://www.bniconnectglobal.com/login/',
  SEARCH_URL: 'https://www.bniconnectglobal.com/web/dashboard/search',

  BASE_DIR:             __dirname,
  CATEGORY_COUNTRY_XLS: path.join(__dirname, 'data', 'Category_Country.xlsx'),
  PROFILE_DIR:          path.join(__dirname, 'output', 'BNIPROFILE'),
  MOVED_DIR:            path.join(__dirname, 'output', 'BNIPROFILE', 'MOVED_FILES'),
  DATA_TXT:             path.join(__dirname, 'data', 'BNI_DATA.txt'),

  NAV_TIMEOUT:       60_000,
  ACTION_TIMEOUT:    20_000,
  WAIT_AFTER_LOGIN:   6_000,
  WAIT_PROFILE_LOAD:  4_000,
  WAIT_AFTER_SEARCH:  5_000,

  // How many End keypresses to scroll-load all lazy results
  PAGE_DOWN_COUNT: 500,

  BROWSER_CHANNEL: 'chrome',
  HEADLESS:false,
};

export default cfg;