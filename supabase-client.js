/**
 * HawkSpot · supabase-client.js
 * Initializes the Supabase client.
 * Replace YOUR_SUPABASE_URL and YOUR_SUPABASE_ANON_KEY in the
 * HAWKSPOT_CONFIG block on each HTML page before deploying.
 *
 * The supabase-js v2 library is loaded via CDN in each HTML page.
 * This module reads config from window.HAWKSPOT_CONFIG so credentials
 * are defined in one place per page and not hard-coded here.
 */

(function () {
    const cfg = window.HAWKSPOT_CONFIG || {};
    const url = cfg.SUPABASE_URL || '';
    const key = cfg.SUPABASE_KEY || '';

    // Expose a supabaseClient on window for use by other scripts.
    // If credentials are placeholders or missing, supabaseClient will be null.
    if (
        url &&
        key &&
        url !== 'YOUR_SUPABASE_URL' &&
        key !== 'YOUR_SUPABASE_ANON_KEY' &&
        typeof window.supabase !== 'undefined'
    ) {
        window.supabaseClient = window.supabase.createClient(url, key);
        console.log('[HawkSpot] Supabase client initialized.');
    } else {
        window.supabaseClient = null;
        console.warn(
            '[HawkSpot] Supabase credentials not set. ' +
            'Running in DEMO MODE with mock data. ' +
            'Set SUPABASE_URL and SUPABASE_KEY in window.HAWKSPOT_CONFIG to connect.'
        );
    }
})();
