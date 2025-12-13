let allCookies = [];
let filteredCookies = [];
let currentEditingCookie = null;
let m365TokenData = null;
let m365Sessions = [];
let activeM365Session = null;

const ROWS_PER_PAGE = 50;
let currentPage = 0;
let searchDebounceTimer = null;
let isPopupMode = true;
let isConvertingToWindow = false;

const TOKEN_STORAGE_KEY = "m365_tokens";
const SESSIONS_STORAGE_KEY = "m365_sessions";
const THEME_STORAGE_KEY = "session_sushi_theme";
const UI_STATE_STORAGE_KEY = "session_sushi_ui_state";
const AUTO_REFRESH_STORAGE_KEY = "m365_auto_refresh_enabled";

let scrollListener = null;
let isLoadingMore = false;
let authWindowId = null;
let autoRefreshTimer = null;
