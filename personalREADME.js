// ==UserScript==
// @name         Roomba Scraper OSINT Profile XLSX Collector v2
// @namespace    https://falconai.local/osint
// @version      2.6.0
// @description  Captura OSINT perfiles XLSX. Nombre LISTADO_CUENTAS_FRAUDULENTAS_CSMR_PLATAFORMA_DD-MM-YYYY. Correcciones FB (Seguidos, ID Comet/Page, Fecha creación en modal y transparencia, sin columna ubicación) y TikTok (Detección real de existencia y ubicación por vídeo/embed).
// @author       Roomba Scraper
// @match        https://www.tiktok.com/@*
// @match        https://x.com/*
// @match        https://twitter.com/*
// @match        https://www.facebook.com/*
// @match        https://facebook.com/*
// @match        https://m.facebook.com/*
// @match        https://www.instagram.com/*
// @match        https://www.threads.com/@*
// @match        https://threads.com/@*
// @match        https://www.threads.net/@*
// @match        https://threads.net/@*
// @require      https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_deleteValue
// @grant        GM_addStyle
// @run-at       document-idle
// ==/UserScript==


(() => {
    'use strict';


    /********************************************************************
     * CONFIGURACIÓN
     ********************************************************************/


    const RECORD_MODE = 'upsert';


    // Pausa entre perfiles de una cola automática.
    const AUTO_MIN_SECONDS = 12;
    const AUTO_MAX_SECONDS = 46;


    const HEADERS = [
        'URL',
        'Motivación del perfil',
        'USUARIO',
        'NOMBRE VISIBLE',
        'ID interno',
        'Nº SEGUIDORES',
        'Nº SEGUIDOS',
        'TIPO DE OBJETO',
        'PLATAFORMA',
        'FECHA DE CREACIÓN DE CUENTA',
        'UBICACIÓN DE LA CUENTA',
        'ÚLTIMO CAMBIO DE USUARIO',
        'ÚLTIMA ACTUALIZACIÓN DEL PERFIL',
        'Nº CAMBIOS DE USUARIO',
        'EXISTE',
        'OBSERVACIONES',
        'GRUPO PÚBLICO',
        'GRUPO VISIBLE'
    ];


    const COLUMN_WIDTHS = {
        'URL': 48,
        'Motivación del perfil': 36,
        'USUARIO': 28,
        'NOMBRE VISIBLE': 34,
        'ID interno': 25,
        'Nº SEGUIDORES': 18,
        'Nº SEGUIDOS': 18,
        'TIPO DE OBJETO': 18,
        'PLATAFORMA': 15,
        'FECHA DE CREACIÓN DE CUENTA': 30,
        'UBICACIÓN DE LA CUENTA': 28,
        'ÚLTIMO CAMBIO DE USUARIO': 28,
        'ÚLTIMA ACTUALIZACIÓN DEL PERFIL': 32,
        'Nº CAMBIOS DE USUARIO': 25,
        'EXISTE': 12,
        'OBSERVACIONES': 70,
        'GRUPO PÚBLICO': 17,
        'GRUPO VISIBLE': 17
    };


    const PLATFORM_HEADER_LABELS = {
        Facebook: {
            'NOMBRE VISIBLE': 'FB: NOMBRE VISIBLE',
            'ID interno': 'FB: ID interno',
            'Nº SEGUIDORES': 'FB: AMIGOS/MIEMBROS',
            'Nº SEGUIDOS': 'FB: SEGUIDOS',
            'FECHA DE CREACIÓN DE CUENTA': 'FB: FECHA DE CREACIÓN DE CUENTA',
            'ÚLTIMA ACTUALIZACIÓN DEL PERFIL': 'FB: ÚLTIMA ACTUALIZACIÓN DEL PERFIL',
            'GRUPO PÚBLICO': 'FB: GRUPO PÚBLICO',
            'GRUPO VISIBLE': 'FB: GRUPO VISIBLE'
        },
        'Twitter/X': {
            'ID interno': 'X: ID interno',
            'FECHA DE CREACIÓN DE CUENTA': 'X: FECHA DE CREACIÓN DE CUENTA',
            'UBICACIÓN DE LA CUENTA': 'X: UBICACIÓN DE LA CUENTA',
            'ÚLTIMO CAMBIO DE USUARIO': 'X: ÚLTIMO CAMBIO DE USUARIO',
            'Nº CAMBIOS DE USUARIO': 'X: Nº CAMBIOS DE USUARIO'
        },
        Instagram: {
            'ID interno': 'IG: ID interno',
            'FECHA DE CREACIÓN DE CUENTA': 'IG: FECHA DE CREACIÓN DE CUENTA',
            'UBICACIÓN DE LA CUENTA': 'IG: UBICACIÓN DE LA CUENTA',
            'Nº CAMBIOS DE USUARIO': 'IG: Nº CAMBIOS DE USUARIO'
        },
        TikTok: {
            'ID interno': 'TT: ID interno',
            'FECHA DE CREACIÓN DE CUENTA': 'TT: FECHA DE CREACIÓN DE CUENTA',
            'UBICACIÓN DE LA CUENTA': 'TT: UBICACIÓN DE LA CUENTA',
            'ÚLTIMO CAMBIO DE USUARIO': 'TT: ÚLTIMO CAMBIO DE NICKNAME'
        },
        Threads: {
            'ID interno': 'TH: ID interno'
        }
    };


    const PLATFORM = detectPlatform();
    if (!PLATFORM) return;


    const SAFE_PLATFORM = PLATFORM
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, '_');


    const DATA_KEY = `falconai_${SAFE_PLATFORM}_rows_v2`;
    const QUEUE_KEY = `falconai_${SAFE_PLATFORM}_queue_v2`;
    const PENDING_KEY = `falconai_${SAFE_PLATFORM}_pending_v2`;


    const PLATFORM_FILE_NAMES = {
        'TikTok':    'TIKTOK',
        'Twitter/X': 'TWITTER',
        'Facebook':  'FACEBOOK',
        'Instagram': 'INSTAGRAM',
        'Threads':   'THREADS'
    };


    function getPlatformHeaders(platform = PLATFORM) {
        if (platform === 'Facebook') {
            return HEADERS.filter(header => header !== 'UBICACIÓN DE LA CUENTA');
        }
        return [...HEADERS];
    }


    function buildXlsxName() {
        const now = new Date();
        const dd   = String(now.getDate()).padStart(2, '0');
        const mm   = String(now.getMonth() + 1).padStart(2, '0');
        const yyyy = now.getFullYear();
        const platformLabel = PLATFORM_FILE_NAMES[PLATFORM] || SAFE_PLATFORM.toUpperCase();
        return `LISTADO_CUENTAS_FRAUDULENTAS_CSMR_${platformLabel}_${dd}-${mm}-${yyyy}.xlsx`;
    }


    GM_deleteValue(`falconai_${SAFE_PLATFORM}_debug_v2`);


    let captureRunning = false;
    let instagramAboutFailure = '';


    /********************************************************************
     * ESTILOS
     ********************************************************************/


    GM_addStyle(`
        #falconai-panel {
            position: fixed;
            right: 14px;
            bottom: 14px;
            width: 365px;
            max-height: 92vh;
            overflow-y: auto;
            z-index: 2147483647;
            background: #0b1220;
            color: #f4f7fb;
            font-family: Arial, sans-serif;
            font-size: 12px;
            border: 1px solid #334155;
            border-radius: 12px;
            box-shadow: 0 8px 30px rgba(0,0,0,.40);
            padding: 12px;
        }


        #falconai-panel h3 {
            margin: 0 0 8px;
            font-size: 14px;
            color: #7dd3fc;
        }


        #falconai-panel button {
            margin: 3px 2px;
            padding: 7px 8px;
            border: 0;
            border-radius: 7px;
            cursor: pointer;
            background: #2563eb;
            color: #fff;
            font-size: 12px;
        }


        #falconai-panel button.secondary {
            background: #475569;
        }


        #falconai-panel button.warn {
            background: #b45309;
        }


        #falconai-panel button.danger {
            background: #991b1b;
        }


        #falconai-panel textarea {
            width: 100%;
            min-height: 85px;
            margin-top: 7px;
            box-sizing: border-box;
            background: #111827;
            color: #f9fafb;
            border: 1px solid #334155;
            border-radius: 8px;
            padding: 7px;
            resize: vertical;
        }


        #falconai-status {
            margin-top: 8px;
            color: #cbd5e1;
            line-height: 1.4;
            white-space: pre-wrap;
        }


        #falconai-small {
            color: #94a3b8;
            margin-top: 6px;
            font-size: 11px;
            line-height: 1.35;
        }
    `);


    /********************************************************************
     * INICIO
     ********************************************************************/


    boot().catch(error => {
        console.error('[Roomba Scraper] Error de inicio:', error);
    });


    async function boot() {
        renderPanel();
        await sleep(1800);


        const resumed = await resumePendingCapture();


        if (!resumed) {
            setTimeout(runAutoQueueSafely, 1000);
        }
    }


    /********************************************************************
     * PANEL
     ********************************************************************/


    function renderPanel() {
        if (document.getElementById('falconai-panel')) return;
        if (!document.body) return;


        const panel = document.createElement('div');
        panel.id = 'falconai-panel';


        panel.innerHTML = `
            <h3>🧹 Roomba Scraper · ${escapeHtml(PLATFORM)}</h3>


            <div>
                <button id="falconai-capture">Capturar perfil</button>
                <button id="falconai-export">Exportar XLSX</button>
                <button id="falconai-import" class="secondary">Importar XLSX</button>
            </div>


            <div>
                <button id="falconai-save-queue" class="secondary">Guardar cola</button>
                <button id="falconai-start-queue" class="warn">Iniciar cola</button>
                <button id="falconai-pause-queue" class="secondary">Pausar</button>
                <button id="falconai-clear" class="danger">Borrar datos</button>
            </div>


            <textarea id="falconai-queue-input"
                placeholder="Pega usuarios o URLs, uno por línea.


usuario1
@usuario2
https://x.com/usuario3"></textarea>


            <input
                id="falconai-file"
                type="file"
                accept=".xlsx,.xls"
                style="display:none"
            >


            <div id="falconai-status">
                Listo. Plataforma detectada: ${escapeHtml(PLATFORM)}
            </div>


            <div id="falconai-small"></div>
        `;


        document.body.appendChild(panel);


        $('#falconai-capture').addEventListener('click', async () => {
            await captureCurrentProfile(false);
        });


        $('#falconai-export').addEventListener('click', exportXlsx);


        $('#falconai-import').addEventListener('click', () => {
            $('#falconai-file').click();
        });


        $('#falconai-file').addEventListener('change', async event => {
            const file = event.target.files?.[0];


            if (file) {
                await importXlsx(file);
            }


            event.target.value = '';
        });


        $('#falconai-save-queue').addEventListener(
            'click',
            saveQueueFromTextarea
        );


        $('#falconai-start-queue').addEventListener('click', () => {
            saveQueueFromTextarea(false);
            const queue = getQueue();


            if (!queue.items.length) {
                setStatus('La cola está vacía.');
                return;
            }


            queue.active = true;


            if (queue.index >= queue.items.length) {
                queue.index = 0;
            }


            saveQueue(queue);
            setStatus(`Cola iniciada. Perfiles: ${queue.items.length}`);
            runAutoQueueSafely();
        });


        $('#falconai-pause-queue').addEventListener('click', () => {
            const queue = getQueue();
            queue.active = false;
            saveQueue(queue);
            setStatus('Cola pausada.');
            updatePanelCounts();
        });


        $('#falconai-clear').addEventListener('click', () => {
            const confirmed = confirm(
                `¿Borrar los datos locales de ${PLATFORM}?`
            );


            if (!confirmed) return;


            GM_deleteValue(DATA_KEY);
            GM_deleteValue(QUEUE_KEY);
            GM_deleteValue(PENDING_KEY);


            setStatus('Datos locales, cola y captura pendiente borrados.');
            refreshQueueTextarea();
            updatePanelCounts();
        });


        refreshQueueTextarea();
        updatePanelCounts();
    }


    /********************************************************************
     * CONTROL DE CAPTURA
     ********************************************************************/


    async function captureCurrentProfile(fromQueue = false) {
        if (captureRunning) {
            setStatus('Ya hay una captura en curso.');
            return null;
        }


        captureRunning = true;


        try {
            await waitForPage();


            if (detectBlock()) {
                pauseQueue();


                return await commitRow(
                    buildUnavailableRow(
                        getUserFromUrl() || 'N/D',
                        'BLOQUEO',
                        'La plataforma ha mostrado una verificación, límite o bloqueo temporal.'
                    ),
                    false
                );
            }


            if (detectNotFound()) {
                if (
                    PLATFORM === 'Facebook' &&
                    startFacebookNumericIdFallback(fromQueue)
                ) {
                    return null;
                }


                return await commitRow(
                    buildUnavailableRow(
                        getUserFromUrl() || 'N/D',
                        'NO',
                        'Perfil, página o grupo no encontrado o no accesible.'
                    ),
                    fromQueue
                );
            }


            switch (PLATFORM) {
                case 'TikTok':
                    return await commitRow(
                        await extractTikTok(),
                        fromQueue
                    );


                case 'Twitter/X':
                    return await captureXFlow(fromQueue);


                case 'Instagram':
                    return await captureInstagramFlow(fromQueue);


                case 'Facebook':
                    return await captureFacebookFlow(fromQueue);


                case 'Threads':
                    return await commitRow(
                        extractThreads(),
                        fromQueue
                    );


                default:
                    return null;
            }
        } catch (error) {
            console.error('[Roomba Scraper] Error de captura:', error);


            if (fromQueue) {
                pauseQueue();
            }


            if (PLATFORM === 'Instagram') {
                const failedRow = extractInstagramBase();
                failedRow.OBSERVACIONES = joinNotes(
                    failedRow.OBSERVACIONES,
                    `Error durante el flujo de Instagram: ${error.message || error}`
                );


                return await commitRow(failedRow, fromQueue);
            }


            setStatus(
                `Error capturando el perfil:\n${error.message || error}`
            );


            return null;
        } finally {
            captureRunning = false;
        }
    }


    async function commitRow(rawData, fromQueue) {
        const replaceFields = Array.isArray(rawData?.__replaceFields)
            ? rawData.__replaceFields
            : [];
        const row = normalizeRow(rawData);
        const result = upsertOrAppend(row, replaceFields);


        setStatus(
            `${result.action}. Filas guardadas: ${result.total}\n` +
            `Usuario: ${row.USUARIO}\n` +
            `ID: ${row['ID interno']}\n` +
            `Tipo: ${row['TIPO DE OBJETO']}\n` +
            `Seguidores/Amigos: ${row['Nº SEGUIDORES']} | Seguidos: ${row['Nº SEGUIDOS']}\n` +
            `Creación: ${row['FECHA DE CREACIÓN DE CUENTA']} | Existe: ${row.EXISTE}`
        );


        updatePanelCounts();


        if (fromQueue) {
            await advanceQueueAfterCommit(row);
        }


        return row;
    }


    /********************************************************************
     * CAPTURA EN DOS FASES: X
     ********************************************************************/


    async function captureXFlow(fromQueue) {
        const username = getUserFromUrl();


        if (!username) {
            return commitRow(
                buildUnavailableRow(
                    'N/D',
                    'NO',
                    'La URL actual no corresponde a un perfil válido de X.'
                ),
                fromQueue
            );
        }


        const existing = findExistingRow(username, 'PERFIL');
        const currentData = extractXPage();
        const combinedBase = mergeRows(existing, currentData);


        if (!isXAboutPage()) {
            savePending({
                stage: 'x-about',
                platform: PLATFORM,
                identity: username,
                fromQueue,
                base: combinedBase
            });


            setStatus(
                `Perfil base capturado.\n` +
                `Abriendo la información principal de:\n` +
                `https://x.com/${username}/about`
            );


            location.href = `https://x.com/${encodeURIComponent(username)}/about`;
            return null;
        }


        const aboutData = await extractXAboutPageWithHistory();


        return commitRow(
            mergeXAboutData(combinedBase, aboutData),
            fromQueue
        );
    }


    function extractXPage() {
        const username = getUserFromUrl() || 'N/D';
        const body = getBodyText();


        return {
            USUARIO: username,
            'ID interno': findXInternalId(username),
            'Nº SEGUIDORES': extractMetric(body, [
                /([\d][\d.,\s]*(?:K|M|B|mil|millones)?)\s+(?:Followers|Seguidores)\b/i,
                /(?:Followers|Seguidores)\s+([\d][\d.,\s]*(?:K|M|B|mil|millones)?)/i
            ]),
            'Nº SEGUIDOS': extractFollowingMetric(body),
            'TIPO DE OBJETO': 'PERFIL',
            MÉTRICA: 'SEGUIDORES',
            PLATAFORMA: PLATFORM,
            'FECHA DE CREACIÓN DE CUENTA': 'N/D',
            'UBICACIÓN DE LA CUENTA': 'N/D',
            'ÚLTIMO CAMBIO DE USUARIO': 'N/D',
            'Nº CAMBIOS DE USUARIO': 'N/D',
            EXISTE: 'SI',
            URL: canonicalProfileUrl(),
            OBSERVACIONES:
                'La fecha de creación y la ubicación se leen exclusivamente desde /about.'
        };
    }


    function extractXAboutPage() {
        const username = getUserFromUrl() || 'N/D';
        const body = getBodyText();


        const joined = extractXJoinedDate(body);


        let accountLocation = rxFirst(body, [
            /Account based in\s*:?\s*([^\n]{2,100})/i,
            /Based in\s*:?\s*([^\n]{2,100})/i,
            /Cuenta basada en\s*:?\s*([^\n]{2,100})/i,
            /Cuenta con sede en\s*:?\s*([^\n]{2,100})/i,
            /La cuenta está ubicada en\s*:?\s*([^\n]{2,100})/i,
            /País o región de la cuenta\s*:?\s*([^\n]{2,100})/i
        ]);


        if (!accountLocation) {
            accountLocation = extractLabeledValue(body, [
                /Account based in/i,
                /Country or region/i,
                /Cuenta basada en/i,
                /Cuenta con sede en/i,
                /País o región/i,
                /Ubicación de la cuenta/i
            ]);
        }


        let connectedVia = rxFirst(body, [
            /Connected via\s*:?\s*([^\n]{2,120})/i,
            /Conectado vía\s*:?\s*([^\n]{2,120})/i
        ]);


        if (!connectedVia) {
            connectedVia = extractLabeledValue(body, [
                /Connected via/i,
                /Conectado vía/i
            ]);
        }


        return {
            USUARIO: username,
            'ID interno': findXInternalId(username),
            'TIPO DE OBJETO': 'PERFIL',
            MÉTRICA: 'SEGUIDORES',
            PLATAFORMA: PLATFORM,
            'FECHA DE CREACIÓN DE CUENTA': joined,
            'UBICACIÓN DE LA CUENTA': accountLocation || 'N/D',
            'ÚLTIMO CAMBIO DE USUARIO': 'N/D',
            'Nº CAMBIOS DE USUARIO': 'N/D',
            EXISTE: 'SI',
            URL: canonicalProfileUrl(),
            OBSERVACIONES: joinNotes(
                'Fecha de creación y ubicación leídas exclusivamente desde /about.',
                connectedVia
                    ? `Conectado vía: ${connectedVia}.`
                    : ''
            )
        };
    }


    async function extractXAboutPageWithHistory() {
        const aboutData = extractXAboutPage();
        const historyData = await openAndReadXUsernameHistory();


        return mergeRows(aboutData, historyData);
    }


    async function openAndReadXUsernameHistory() {
        const labels = [
            'Username history',
            'Username changes',
            'Historial de nombres de usuario',
            'Historial del nombre de usuario',
            'Cambios de nombre de usuario'
        ];


        let history = parseXUsernameHistory(getBodyText());
        const historyEntry = findClickableByText(labels);


        if (historyEntry) {
            safeClick(historyEntry);


            await waitFor(() => {
                const content = normalizeText(getBodyText());


                return /last changed|most recent change|ultimo cambio|cambio mas reciente|no former usernames|ningun nombre de usuario anterior|username changes|cambios de nombre de usuario/.test(content);
            }, 6000, 250);


            history = parseXUsernameHistory(getBodyText());
        }


        return history;
    }


    function parseXUsernameHistory(source) {
        const content = String(source || '');
        const normalized = normalizeText(content);


        let changeCount = rxFirst(content, [
            /(?:Username changes?|Former usernames?)\s*:?\s*(\d{1,5})/i,
            /(?:Username history)\s*:?\s*(\d{1,5})/i,
            /(?:Username history)[\s\S]{0,100}?\b(\d{1,5})\s+(?:changes?|times?)/i,
            /(\d{1,5})\s+(?:username changes?|former usernames?)/i,
            /(?:Número de cambios de usuario|Cambios de nombre de usuario|Nombres de usuario anteriores)\s*:?\s*(\d{1,5})/i,
            /(?:Historial de nombres? de usuario)\s*:?\s*(\d{1,5})/i,
            /(?:Historial de nombres? de usuario)[\s\S]{0,100}?\b(\d{1,5})\s+(?:cambios?|veces?)/i,
            /(\d{1,5})\s+(?:cambios? de nombre de usuario|nombres? de usuario anteriores?)/i
        ]);


        const noPreviousUsername =
            /no former usernames|no username changes|has never changed (?:their|its) username|ningun nombre de usuario anterior|sin nombres de usuario anteriores|no hay informacion anterior sobre el nombre de usuario|nunca ha cambiado (?:su|el) nombre de usuario/.test(normalized);


        if (!changeCount && noPreviousUsername) {
            changeCount = '0';
        }


        let lastChangeRaw = rxFirst(content, [
            /(?:Last changed|Most recent change)\s*:?\s*([^\n]{3,80})/i,
            /(?:Último cambio|Cambio más reciente|Nombre de usuario cambiado el)\s*:?\s*([^\n]{3,80})/i
        ]);


        if (!lastChangeRaw) {
            lastChangeRaw = extractLabeledValue(content, [
                /Last changed/i,
                /Most recent change/i,
                /Último cambio/i,
                /Cambio más reciente/i,
                /Nombre de usuario cambiado el/i
            ]);
        }


        const lastChange = changeCount === '0'
            ? '0'
            : formatMonthYearEs(lastChangeRaw);


        return {
            'ÚLTIMO CAMBIO DE USUARIO': lastChange,
            'Nº CAMBIOS DE USUARIO': changeCount || 'N/D'
        };
    }


    function mergeXAboutData(base, about) {
        const replaceFields = [
            'FECHA DE CREACIÓN DE CUENTA',
            'UBICACIÓN DE LA CUENTA',
            'ÚLTIMO CAMBIO DE USUARIO',
            'Nº CAMBIOS DE USUARIO'
        ];


        const merged = mergeRows(base, about);


        for (const field of replaceFields) {
            merged[field] = about[field] || 'N/D';
        }


        merged.__replaceFields = replaceFields;
        return merged;
    }


    function extractXJoinedDate(body) {
        const raw = rxFirst(body, [
            /Joined\s*:?\s*([A-Za-z]+\s+\d{4})/i,
            /Joined\s*:?\s*([A-Za-z]+\s+\d{1,2},?\s+\d{4})/i,
            /Se unió el\s*:?\s*([^\n]{3,60})/i,
            /Se unió en\s+([A-Za-zÁÉÍÓÚÜÑáéíóúüñ]+\s+de\s+\d{4})/i,
            /Se unió\s+([A-Za-zÁÉÍÓÚÜÑáéíóúüñ]+\s+\d{4})/i,
            /Fecha de incorporación\s*:?\s*([^\n]{3,60})/i
        ]);


        return formatMonthYearEs(raw);
    }


    /********************************************************************
     * INSTAGRAM: POP-UP "ACERCA DE ESTA CUENTA"
     ********************************************************************/


    async function captureInstagramFlow(fromQueue) {
        const base = extractInstagramBase();


        setStatus(
            'Buscando y abriendo “Acerca de esta cuenta”…'
        );


        const about = await openAndReadInstagramAbout();


        if (!about) {
            base.OBSERVACIONES = joinNotes(
                base.OBSERVACIONES,
                instagramAboutFailure ||
                    'No se pudo abrir “Acerca de esta cuenta”. Puede no estar disponible para ese perfil, sesión o versión de interfaz.'
            );


            return commitRow(base, fromQueue);
        }


        return commitRow(
            mergeRows(base, about),
            fromQueue
        );
    }


    function extractInstagramBase() {
        const username = getUserFromUrl() || 'N/D';


        const description =
            metaContent('meta[property="og:description"]') ||
            metaContent('meta[name="description"]') ||
            '';


        const followers =
            rxFirst(description, [
                /^([^,]+)\s+(?:Followers|seguidores)/i,
                /([\d][\d.,\s]*(?:K|M|B|mil|millones)?)\s+(?:Followers|seguidores)/i
            ]) ||
            extractMetric(getBodyText(), [
                /([\d][\d.,\s]*(?:K|M|B|mil|millones)?)\s+(?:Followers|seguidores)\b/i
            ]);


        const following =
            extractFollowingMetric(description) ||
            extractFollowingMetric(getBodyText());


        return {
            USUARIO: username,
            'ID interno': findInstagramInternalId(username),
            'Nº SEGUIDORES': followers || 'N/D',
            'Nº SEGUIDOS': following || 'N/D',
            'TIPO DE OBJETO': 'PERFIL',
            MÉTRICA: 'SEGUIDORES',
            PLATAFORMA: PLATFORM,
            'FECHA DE CREACIÓN DE CUENTA': 'N/D',
            'UBICACIÓN DE LA CUENTA': 'N/D',
            'ÚLTIMO CAMBIO DE USUARIO': 'N/D',
            'Nº CAMBIOS DE USUARIO': 'N/D',
            EXISTE: 'SI',
            URL: canonicalProfileUrl(),
            OBSERVACIONES:
                'El ID se busca únicamente en datos embebidos asociados al nombre de usuario.'
        };
    }


    async function openAndReadInstagramAbout() {
        instagramAboutFailure = '';
        let dialog = findInstagramAboutDialog();


        if (dialog) {
            return parseInstagramAboutDialog(dialog);
        }


        const optionsButton = await waitForDom(
            findInstagramOptionsButton,
            10000
        );


        const optionsClicked = optionsButton
            ? safeClick(optionsButton)
            : false;


        if (!optionsButton || !optionsClicked) {
            instagramAboutFailure =
                'Instagram: no se encontró o no se pudo activar el botón de tres puntos.';
            return null;
        }


        const menuItem = await waitForDom(
            findInstagramAboutMenuItem,
            8000
        );


        const menuItemClicked = menuItem
            ? safeClick(menuItem)
            : false;


        if (!menuItem || !menuItemClicked) {
            instagramAboutFailure =
                'Instagram: el menú se abrió, pero no se encontró o no se pudo activar “Información sobre esta cuenta”.';
            return null;
        }


        dialog = await waitForDom(
            findInstagramAboutDialog,
            12000
        );


        if (!dialog) {
            instagramAboutFailure =
                'Instagram: se activó “Información sobre esta cuenta”, pero el popup con los datos no apareció.';
            return null;
        }


        return parseInstagramAboutDialog(dialog);
    }


    function findInstagramOptionsButton() {
        const labels = [
            'options',
            'more options',
            'opciones',
            'más opciones',
            'mas opciones'
        ];


        const scopes = [
            document.querySelector('main header'),
            document.querySelector('header'),
            document.querySelector('main'),
            document
        ].filter(Boolean);


        for (const scope of scopes) {
            const popupButtons = Array.from(scope.querySelectorAll(
                '[role="button"][aria-haspopup="dialog"]'
            ));


            for (const button of popupButtons) {
                const accessibleText = normalizeText([
                    button.getAttribute('aria-label') || '',
                    button.getAttribute('title') || '',
                    button.querySelector('[aria-label]')
                        ?.getAttribute('aria-label') || '',
                    button.querySelector('title')?.textContent || ''
                ].join(' '));


                if (
                    labels.some(label => {
                        return accessibleText.includes(
                            normalizeText(label)
                        );
                    }) &&
                    isVisible(button)
                ) {
                    return button;
                }
            }


            const elements = Array.from(scope.querySelectorAll(
                '[aria-label],[title]'
            ));


            for (const element of elements) {
                const aria = normalizeText(
                    element.getAttribute('aria-label') ||
                    element.getAttribute('title') ||
                    ''
                );


                if (!labels.some(label => aria === normalizeText(label))) {
                    continue;
                }


                const clickable =
                    element.closest(
                        'button,[role="button"],div[tabindex="0"]'
                    ) ||
                    element;


                if (isVisible(clickable)) {
                    return clickable;
                }
            }
        }


        return null;
    }


    function findInstagramAboutMenuItem() {
        const labels = [
            'About this account',
            'Acerca de esta cuenta',
            'Información sobre esta cuenta'
        ].map(normalizeText);


        const textElements = Array.from(document.querySelectorAll(
            'span,div[role="button"],[role="menuitem"]'
        ));


        for (const element of textElements) {
            if (normalizeText(element.textContent || '') === '') {
                continue;
            }


            const clickable = element.closest(
                '[role="button"],[role="menuitem"],button'
            );


            if (!clickable || !isVisible(clickable)) continue;


            const value = normalizeText(
                clickable.innerText ||
                clickable.textContent ||
                ''
            );


            if (labels.includes(value)) {
                return clickable;
            }
        }


        return null;
    }


    function findInstagramAboutDialog() {
        const dialogs = Array.from(
            document.querySelectorAll(
                '[role="dialog"][aria-modal="true"], [role="dialog"]'
            )
        ).reverse();


        for (const dialog of dialogs) {
            const content = normalizeText(
                dialog.innerText || dialog.textContent || ''
            );


            if (
                /date joined|fecha de registro|fecha en que se unio|fecha en la que te uniste|se unio a instagram|account based in|cuenta basada en|ubicacion de la cuenta|pais donde se encuentra la cuenta|former username|nombres de usuario anteriores/.test(content)
            ) {
                return dialog;
            }
        }


        return null;
    }


    function parseInstagramAboutDialog(dialog) {
        const content =
            dialog.innerText ||
            dialog.textContent ||
            '';


        let dateRaw = extractInstagramDialogValue(dialog, [
            'Date joined',
            'Fecha de registro',
            'Fecha en que se unió',
            'Fecha en la que te uniste',
            'Se unió a Instagram'
        ]);


        if (!dateRaw) {
            dateRaw = rxFirst(content, [
                /Date joined\s*:?\s*([A-Za-z]+\s+\d{4})/i,
                /Joined\s*:?\s*([A-Za-z]+\s+\d{4})/i,
                /Fecha de registro\s*:?\s*([^\n]{3,70})/i,
                /Fecha en que se unió\s*:?\s*([^\n]{3,70})/i,
                /Fecha en la que te uniste\s*:?\s*([^\n]{3,70})/i,
                /Se unió a Instagram en\s*:?\s*([^\n]{3,70})/i
            ]);
        }


        if (!dateRaw) {
            dateRaw = extractLabeledValue(content, [
                /Date joined/i,
                /Fecha de registro/i,
                /Fecha en que se unió/i,
                /Fecha en la que te uniste/i,
                /Se unió a Instagram/i
            ]);
        }


        let accountLocation = extractInstagramDialogValue(dialog, [
            'Account based in',
            'Cuenta basada en',
            'País donde se encuentra la cuenta',
            'Ubicación de la cuenta'
        ]);


        if (!accountLocation) {
            accountLocation = rxFirst(content, [
                /Account based in\s*:?\s*([^\n]{2,100})/i,
                /Cuenta basada en\s*:?\s*([^\n]{2,100})/i,
                /País donde se encuentra la cuenta\s*:?\s*([^\n]{2,100})/i,
                /Ubicación de la cuenta\s*:?\s*([^\n]{2,100})/i
            ]);
        }


        if (!accountLocation) {
            accountLocation = extractLabeledValue(content, [
                /Account based in/i,
                /Cuenta basada en/i,
                /País donde se encuentra la cuenta/i,
                /Ubicación de la cuenta/i
            ]);
        }


        const changeCount = extractCountAroundLabel(content, [
            /Former username count/i,
            /Former usernames/i,
            /Username changes/i,
            /Número de nombres de usuario anteriores/i,
            /Nombres de usuario anteriores/i,
            /Cambios de nombre de usuario/i
        ]);


        return {
            USUARIO: getUserFromUrl() || 'N/D',
            'ID interno': findInstagramInternalId(
                getUserFromUrl()
            ),
            'TIPO DE OBJETO': 'PERFIL',
            MÉTRICA: 'SEGUIDORES',
            PLATAFORMA: PLATFORM,
            'FECHA DE CREACIÓN DE CUENTA':
                formatMonthYearEs(dateRaw),
            'UBICACIÓN DE LA CUENTA':
                cleanAccountLocation(accountLocation),
            'ÚLTIMO CAMBIO DE USUARIO': 'N/D',
            'Nº CAMBIOS DE USUARIO':
                changeCount || 'N/D',
            EXISTE: 'SI',
            URL: canonicalProfileUrl(),
            OBSERVACIONES:
                'Fecha, país y número de cambios extraídos de “Acerca de esta cuenta”.'
        };
    }


    function extractInstagramDialogValue(dialog, labels) {
        const normalizedLabels = labels.map(normalizeText);
        const candidates = Array.from(dialog.querySelectorAll(
            '[aria-label],[data-bloks-name="bk.components.Flexbox"]'
        ));


        for (const candidate of candidates) {
            const aria = clean(
                candidate.getAttribute('aria-label') || ''
            );


            const lines = String(
                candidate.innerText ||
                candidate.textContent ||
                ''
            )
                .split(/\r?\n/)
                .map(clean)
                .filter(Boolean);


            for (const label of normalizedLabels) {
                const lineIndex = lines.findIndex(line => {
                    return normalizeText(line) === label;
                });


                if (lineIndex >= 0 && lines[lineIndex + 1]) {
                    return lines[lineIndex + 1];
                }


                const normalizedAria = normalizeText(aria);


                if (normalizedAria.startsWith(`${label} `)) {
                    return clean(
                        aria.slice(labels[normalizedLabels.indexOf(label)].length)
                    );
                }
            }
        }


        return '';
    }


    /********************************************************************
     * FACEBOOK: PERFIL, GRUPO Y /ABOUT_PROFILE_TRANSPARENCY
     ********************************************************************/


    function startFacebookNumericIdFallback(fromQueue) {
        const profileId = (() => {
            try {
                const url = new URL(location.href);
                const id = url.searchParams.get('id') || '';


                return (
                    url.pathname.startsWith('/profile.php') &&
                    /^\d{5,}$/.test(id)
                )
                    ? id
                    : '';
            } catch {
                return '';
            }
        })();


        if (!profileId) return false;


        savePending({
            stage: 'facebook-id-fallback',
            platform: PLATFORM,
            identity: profileId,
            fromQueue
        });


        setStatus(
            `El ID ${profileId} no corresponde a un perfil accesible.\n` +
            'Probándolo como grupo de Facebook…'
        );


        setTimeout(() => {
            location.href =
                `https://www.facebook.com/groups/${encodeURIComponent(profileId)}`;
        }, 700);


        return true;
    }


    async function captureFacebookFlow(fromQueue) {
        const type = getFacebookObjectType();
        const identity = getFacebookIdentity();


        if (!identity) {
            return commitRow(
                buildUnavailableRow(
                    'N/D',
                    'NO',
                    'No se pudo identificar el perfil o grupo de Facebook.'
                ),
                fromQueue
            );
        }


        let currentData = extractFacebookPage();


        if (type === 'PERFIL') {
            setStatus('Abriendo información / transparencia del perfil de Facebook…');
            const popupData = await openAndReadFacebookProfileInfo();
            currentData = mergeRows(currentData, popupData);
        }


        const existing = findExistingRow(
            currentData.USUARIO,
            type,
            currentData['ID interno']
        );


        const combinedBase = mergeRows(existing, currentData);


        const needsAboutNavigation =
            !isFacebookAboutPage() &&
            (
                !isUsefulValue(combinedBase['FECHA DE CREACIÓN DE CUENTA']) ||
                (
                    type === 'GRUPO' &&
                    (
                        !isUsefulValue(combinedBase['GRUPO PÚBLICO']) ||
                        !isUsefulValue(combinedBase['GRUPO VISIBLE'])
                    )
                )
            );


        if (needsAboutNavigation) {
            const aboutUrl = facebookAboutUrl();


            if (aboutUrl) {
                savePending({
                    stage: 'facebook-about',
                    platform: PLATFORM,
                    identity,
                    fromQueue,
                    base: combinedBase
                });


                setStatus(
                    `${type === 'GRUPO' ? 'Grupo' : 'Perfil'} base capturado.\n` +
                    `Abriendo sección de transparencia / Acerca de…`
                );


                location.href = aboutUrl;
                return null;
            }
        }


        return commitRow(
            mergeRows(combinedBase, extractFacebookPage()),
            fromQueue
        );
    }


    function extractFacebookPage() {
        const body = getBodyText();
        const type = getFacebookObjectType();


        if (type === 'GRUPO') {
            const groupName = extractFacebookGroupName();
            const groupId = getFacebookGroupIdFromUrl() || getFacebookIdentity();


            return {
                USUARIO: groupId || 'N/D',
                'NOMBRE VISIBLE': groupName || 'N/D',
                'ID interno': findFacebookInternalId('GRUPO'),
                'Nº SEGUIDORES': extractFacebookGroupMembers(body),
                'Nº SEGUIDOS': 'N/D',
                'TIPO DE OBJETO': 'GRUPO',
                MÉTRICA: 'MIEMBROS',
                PLATAFORMA: PLATFORM,
                'FECHA DE CREACIÓN DE CUENTA':
                    extractFacebookCreationDate(body),
                'ÚLTIMO CAMBIO DE USUARIO': 'N/D',
                'ÚLTIMA ACTUALIZACIÓN DEL PERFIL': 'N/D',
                'Nº CAMBIOS DE USUARIO': 'N/D',
                'GRUPO PÚBLICO':
                    extractFacebookGroupPublic(body),
                'GRUPO VISIBLE':
                    extractFacebookGroupVisible(body),
                EXISTE: 'SI',
                URL: canonicalProfileUrl(),
                OBSERVACIONES:
                    'Para grupos, USUARIO contiene el identificador del grupo y Nº SEGUIDORES contiene sus miembros.'
            };
        }


        const urlName = getFacebookProfileUrlName() || 'N/D';
        const displayName = extractFacebookDisplayName() || 'N/D';


        return {
            USUARIO: urlName,
            'NOMBRE VISIBLE': displayName,
            'ID interno': findFacebookInternalId('PERFIL'),
            'Nº SEGUIDORES': extractFacebookFriends(body),
            'Nº SEGUIDOS': extractFacebookFollowing(body),
            'TIPO DE OBJETO': 'PERFIL',
            MÉTRICA: 'AMIGOS',
            PLATAFORMA: PLATFORM,
            'FECHA DE CREACIÓN DE CUENTA':
                extractFacebookCreationDate(body),
            'ÚLTIMO CAMBIO DE USUARIO': 'N/D',
            'ÚLTIMA ACTUALIZACIÓN DEL PERFIL':
                extractFacebookLastUpdated(body) || 'N/D',
            'Nº CAMBIOS DE USUARIO': 'N/D',
            EXISTE: 'SI',
            URL: canonicalProfileUrl(),
            OBSERVACIONES:
                'Datos extraídos de la cabecera, menú de información del perfil y transparencia.'
        };
    }


    async function openAndReadFacebookProfileInfo() {
        let dialog = findFacebookProfileInfoDialog();


        if (!dialog) {
            const optionsButton = findFacebookOptionsButton();


            if (optionsButton) {
                safeClick(optionsButton);


                const menuItem = await waitForDom(
                    findFacebookProfileInfoMenuItem,
                    5000
                );


                if (menuItem) {
                    safeClick(menuItem);


                    dialog = await waitForDom(
                        findFacebookProfileInfoDialog,
                        8000
                    );
                }
            }
        }


        if (!dialog) return null;


        // Esperar a que el contenido interno (GraphQL) del modal termine de renderizar la fecha
        await waitFor(() => {
            const textContent = dialog.innerText || dialog.textContent || '';
            return (
                extractFacebookCreationDate(textContent) !== 'N/D' ||
                /\b(?:19|20)\d{2}\b/.test(textContent)
            );
        }, 5500, 250);


        const data = parseFacebookProfileInfoDialog(dialog);
        closeFacebookDialog(dialog);
        return data;
    }


    function findFacebookOptionsButton() {
        const exactLabels = [
            'acciones para este perfil',
            'acciones del perfil',
            'más opciones',
            'mas opciones',
            'opciones del perfil',
            'ver más opciones',
            'ver mas opciones',
            'profile actions',
            'actions for this profile',
            'more options',
            'profile options',
            'see more options'
        ];


        const scopes = [
            document.querySelector('[role="main"]'),
            document.querySelector('main'),
            document.body
        ].filter(Boolean);


        for (const scope of scopes) {
            const buttons = Array.from(scope.querySelectorAll(
                '[role="button"][aria-label], button[aria-label], [aria-haspopup="menu"][aria-label]'
            ));


            for (const button of buttons) {
                if (!isVisible(button)) continue;


                // No pulsar botones dentro de publicaciones/artículos del muro
                if (button.closest('[role="article"]')) continue;


                const aria = normalizeText(
                    button.getAttribute('aria-label') ||
                    button.getAttribute('title') ||
                    ''
                );


                if (exactLabels.some(label => aria === label || aria.startsWith(label))) {
                    return button;
                }
            }


            // Fallback: botón de menú desplegable con icono SVG de tres puntos en la cabecera del perfil
            const menuButtons = Array.from(scope.querySelectorAll(
                '[role="button"][aria-haspopup="menu"], [role="button"]'
            ));


            for (const button of menuButtons) {
                if (!isVisible(button) || button.closest('[role="article"]') || button.closest('[role="tablist"]')) {
                    continue;
                }


                const circles = button.querySelectorAll('svg circle, svg path');
                const inner = clean(button.innerText || '');
                const aria = normalizeText(button.getAttribute('aria-label') || '');


                if (
                    inner === '' &&
                    circles.length >= 1 &&
                    !/mensaje|message|amigo|friend|seguir|follow|buscar|search|editar|edit|compartir|share|notificaci/i.test(aria) &&
                    button.getAttribute('aria-haspopup') === 'menu'
                ) {
                    return button;
                }
            }
        }


        return null;
    }


    function findFacebookProfileInfoMenuItem() {
        const labels = [
            'informacion sobre este perfil',
            'información sobre este perfil',
            'informacion sobre el perfil',
            'información sobre el perfil',
            'informacion del perfil',
            'información del perfil',
            'informacion sobre esta pagina',
            'información sobre esta página',
            'informacion sobre esta cuenta',
            'información sobre esta cuenta',
            'transparencia del perfil',
            'transparencia de la pagina',
            'transparencia de la página',
            'acerca de este perfil',
            'acerca de esta pagina',
            'acerca de esta cuenta',
            'about this profile',
            'about this page',
            'about this account',
            'profile transparency',
            'page transparency',
            'profile information'
        ].map(normalizeText);


        const candidates = Array.from(document.querySelectorAll(
            '[role="menuitem"], [role="menu"] [role="button"], [role="dialog"] [role="button"], span'
        ));


        for (const element of candidates) {
            const textContent = normalizeText(
                element.innerText || element.textContent || ''
            );


            if (!textContent || textContent.length > 90) continue;


            if (labels.some(label => textContent === label || textContent.includes(label))) {
                const clickable = element.closest(
                    '[role="menuitem"],[role="button"],a,button'
                ) || element;


                if (isVisible(clickable)) {
                    return clickable;
                }
            }
        }


        return null;
    }


    function findFacebookProfileInfoDialog() {
        const dialogs = Array.from(document.querySelectorAll(
            '[role="dialog"], [aria-modal="true"]'
        )).reverse();


        return dialogs.find(dialog => {
            if (!isVisible(dialog)) return false;
            const content = normalizeText(
                dialog.innerText || dialog.textContent || ''
            );


            return /informacion sobre este perfil|informacion sobre el perfil|informacion sobre esta pagina|about this profile|about this page|transparencia del perfil|transparencia de la pagina|profile transparency|page transparency|se unio|fecha de creacion|creacion del perfil|creacion de la pagina|joined facebook|profile created|page created|perfil actualizado/.test(content);
        }) || null;
    }


    function parseFacebookProfileInfoDialog(dialog) {
        const content =
            dialog.innerText ||
            dialog.textContent ||
            '';


        const creation = extractFacebookCreationDate(content);
        const updatedRaw = extractFacebookLastUpdated(content);


        return {
            'FECHA DE CREACIÓN DE CUENTA': creation !== 'N/D' ? creation : 'N/D',
            'ÚLTIMA ACTUALIZACIÓN DEL PERFIL': updatedRaw || 'N/D',
            OBSERVACIONES:
                'Fecha de incorporación y actualización obtenidas del popup del perfil.'
        };
    }


    function extractFacebookLastUpdated(content) {
        let updatedRaw = rxFirst(content, [
            /Perfil actualizado\s*:?\s*([^\n]{3,80})/i,
            /[UÚ]ltima actualizaci[oó]n del perfil\s*:?\s*([^\n]{3,80})/i,
            /[UÚ]ltima actualizaci[oó]n\s*:?\s*([^\n]{3,80})/i,
            /Perfil actualizado el\s*:?\s*([^\n]{3,80})/i,
            /Profile last updated\s*:?\s*([^\n]{3,80})/i,
            /Last profile update\s*:?\s*([^\n]{3,80})/i
        ]);


        if (!updatedRaw) {
            updatedRaw = extractLabeledValue(content, [
                /Perfil actualizado/i,
                /[UÚ]ltima actualizaci[oó]n del perfil/i,
                /[UÚ]ltima actualizaci[oó]n/i,
                /Perfil actualizado el/i,
                /Profile last updated/i,
                /Last profile update/i
            ]);
        }


        return clean(updatedRaw);
    }


    function closeFacebookDialog(dialog) {
        const closeButton = dialog.querySelector(
            '[role="button"][aria-label="Cerrar"], ' +
            '[role="button"][aria-label="Close"]'
        ) || findClickableByText(
            ['Cerrar', 'Close'],
            dialog,
            true
        );


        if (closeButton) {
            safeClick(closeButton);
            return;
        }


        document.dispatchEvent(
            new KeyboardEvent('keydown', {
                key: 'Escape',
                code: 'Escape',
                bubbles: true
            })
        );
    }


    function extractFacebookGroupMembers(body) {
        return extractMetricFromLines(
            body,
            [
                /([\d][\d.,\s]*(?:K|M|B|mil|millones|milhões)?)\s+(?:members|miembros|membros)\b/i,
                /(?:members|miembros|membros)\s*:?\s*([\d][\d.,\s]*(?:K|M|B|mil|millones|milhões)?)/i
            ],
            [
                /online|en línea|en linea|activos ahora/i
            ]
        );
    }


    function extractFacebookGroupPublic(body) {
        const lines = String(body || '')
            .split(/\r?\n/)
            .map(normalizeText)
            .filter(Boolean);


        if (
            lines.some(line => {
                return /^(?:grupo )?publico\b|^public(?: group)?\b|\bgrupo publico\b|\bpublic group\b/.test(line);
            })
        ) {
            return 'SI';
        }


        if (
            lines.some(line => {
                return /^(?:grupo )?privado\b|^private(?: group)?\b|\bgrupo privado\b|\bprivate group\b/.test(line);
            })
        ) {
            return 'NO';
        }


        return 'N/D';
    }


    function extractFacebookGroupVisible(body) {
        const lines = String(body || '')
            .split(/\r?\n/)
            .map(normalizeText)
            .filter(Boolean);


        if (
            lines.some(line => {
                return /^(?:grupo )?visible\b|^visible group\b|\bgrupo visible\b/.test(line);
            })
        ) {
            return 'SI';
        }


        if (
            lines.some(line => {
                return /^(?:grupo )?oculto\b|^hidden(?: group)?\b|\bgrupo oculto\b/.test(line);
            })
        ) {
            return 'NO';
        }


        return 'N/D';
    }


    function extractFacebookFriends(body) {
        // 1. Buscar en enlaces de cabecera (amigos / seguidores)
        const headerLinks = Array.from(document.querySelectorAll(
            '[role="main"] a[href*="/friends"], [role="main"] a[href*="sk=friends"], [role="main"] a[href*="/followers"]'
        ));


        for (const link of headerLinks) {
            const linkText = clean(link.innerText || link.textContent || '');
            if (/mutual|en común|en comun/i.test(linkText)) continue;


            const m = linkText.match(
                /([\d][\d.,\s]*(?:K|M|B|mil|millones|milhões)?)\s*(?:friends|amigos|amigas|followers|seguidores|seguidoras)?\b/i
            );
            if (m?.[1]) return clean(m[1]);
        }


        // 2. Buscar amigos en el texto del perfil
        const friends = extractMetricFromLines(
            body,
            [
                /(?:^|[\s•·|])([\d][\d.,\s]*(?:K|M|B|mil|millones|milhões)?)\s+(?:friends|amigos|amigas)\b/i,
                /\b(?:friends|amigos|amigas)\s*:?\s*([\d][\d.,\s]*(?:K|M|B|mil|millones|milhões)?)/i
            ],
            [
                /mutual|en común|en comun|amigos en común|amigos en comun/i
            ]
        );


        if (friends !== 'N/D') return friends;


        // 3. Fallback para perfiles en Modo Profesional o Páginas que muestran "seguidores" en lugar de amigos
        return extractMetricFromLines(
            body,
            [
                /(?:^|[\s•·|])([\d][\d.,\s]*(?:K|M|B|mil|millones|milhões)?)\s+(?:followers|seguidores|seguidoras)\b/i,
                /(?:Followed by|Seguido por)\s+([\d][\d.,\s]*(?:K|M|B|mil|millones|milhões)?)\s+(?:people|personas)\b/i
            ],
            []
        );
    }


    function extractFacebookFollowing(body) {
        // 1. Buscar directamente en enlaces DOM de "seguidos / following"
        const followingLinks = Array.from(document.querySelectorAll(
            '[role="main"] a[href*="/following"], [role="main"] a[href*="sk=following"]'
        ));


        for (const link of followingLinks) {
            const linkText = clean(link.innerText || link.textContent || '');
            const m = linkText.match(
                /([\d][\d.,\s]*(?:K|M|B|mil|millones|milhões)?)\s*(?:following|siguiendo|seguidos|seguidas)?\b/i
            );
            if (m?.[1]) return clean(m[1]);
        }


        // 2. Buscar en el texto visible (soporta separadores • / · y la frase "Sigue a X personas" de Detalles)
        const fromLines = extractMetricFromLines(
            body,
            [
                /(?:^|[\s•·|,\n\r])([\d][\d.,\s]*(?:K|M|B|mil|millones|milhões)?)\s+(?:following|siguiendo|seguidos|seguidas)\b/i,
                /\b(?:following|siguiendo|seguidos|seguidas)\s*:?\s*([\d][\d.,\s]*(?:K|M|B|mil|millones|milhões)?)/i,
                /\b(?:sigue\s+a)\s+([\d][\d.,\s]*(?:K|M|B|mil|millones|milhões)?)\s+personas?\b/i
            ],
            []
        );


        if (fromLines !== 'N/D') return fromLines;


        // 3. Buscar en los datos JSON / GraphQL embebidos en el HTML
        for (const source of getEmbeddedTextSources()) {
            for (const variant of embeddedSourceVariants(source)) {
                const match = variant.match(
                    /"text"\s*:\s*"([\d][\d.,\s]*(?:K|M|B|mil|millones|milhões)?)\s+(?:following|siguiendo|seguidos|seguidas)"/i
                ) || variant.match(
                    /"text"\s*:\s*"Sigue a\s+([\d][\d.,\s]*(?:K|M|B|mil|millones|milhões)?)\s+personas?"/i
                );


                if (match?.[1]) {
                    return clean(match[1]);
                }
            }
        }


        return 'N/D';
    }


    function extractFacebookCreationDate(body) {
        if (!body) return 'N/D';


        const patterns = [
            /se\s+uni[oó]\s+en\s+([A-Za-zÁÉÍÓÚÜÑáéíóúüñ]+\s+(?:del?\s+)?\d{4})/i,
            /se\s+uni[oó]\s+en\s+(\d{4})/i,
            /se\s+uni[oó](?:\s+a\s+facebook)?\s+el\s+([^\n.,;]{3,80})/i,
            /se\s+uni[oó]\s+a\s+facebook(?:\s+en)?\s+([^\n.,;]{3,80})/i,
            /joined(?:\s+facebook)?(?:\s+in|\s+on)?\s+([A-Za-z]+\s+\d{1,2},?\s+\d{4})/i,
            /joined(?:\s+facebook)?(?:\s+in|\s+on)?\s+([A-Za-z]+\s+\d{4})/i,
            /joined(?:\s+facebook)?(?:\s+in|\s+on)?\s+(\d{4})/i,
            /joined\s+facebook\s*:?\s*([^\n.,;]{3,80})/i,
            /fecha\s+de\s+creaci[oó]n\s+del?\s+perfil\s*:?\s*([^\n.,;]{3,80})/i,
            /fecha\s+de\s+creaci[oó]n\s+de\s+la\s+cuenta\s*:?\s*([^\n.,;]{3,80})/i,
            /fecha\s+de\s+creaci[oó]n\s+de\s+la\s+p[aá]gina\s*:?\s*([^\n.,;]{3,80})/i,
            /fecha\s+de\s+creaci[oó]n\s+del?\s+grupo\s*:?\s*([^\n.,;]{3,80})/i,
            /fecha\s+de\s+creaci[oó]n\s*:?\s*([^\n.,;]{3,80})/i,
            /creaci[oó]n\s+del?\s+(?:perfil|p[aá]gina|cuenta|grupo)\s*:?\s*([^\n.,;]{3,80})/i,
            /(?:perfil|p[aá]gina|cuenta|grupo)\s+cread[oa](?:\s+el|\s+en)?\s*:?\s*([^\n.,;]{3,80})/i,
            /(?:profile|page|account|group)\s+created(?:\s+on|\s+in)?\s*:?\s*([^\n.,;]{3,80})/i,
            /created\s+on\s+([^\n.,;]{3,80})/i,
            /creation\s+date\s*:?\s*([^\n.,;]{3,80})/i,
            /fecha\s+de\s+incorporaci[oó]n(?:\s+a\s+facebook|\s+a\s+la\s+plataforma)?\s*:?\s*([^\n.,;]{3,80})/i,
            /miembro\s+desde\s*:?\s*([^\n.,;]{3,80})/i,
            /member\s+since\s*:?\s*([^\n.,;]{3,80})/i
        ];


        const visibleRaw = rxFirst(body, patterns);


        if (visibleRaw) {
            const formatted = formatMonthYearEs(visibleRaw);
            if (formatted !== 'N/D') return formatted;
        }


        const labels = [
            /fecha\s+de\s+creaci[oó]n\s+del?\s+perfil/i,
            /fecha\s+de\s+creaci[oó]n\s+de\s+la\s+cuenta/i,
            /fecha\s+de\s+creaci[oó]n\s+de\s+la\s+p[aá]gina/i,
            /fecha\s+de\s+creaci[oó]n\s+del?\s+grupo/i,
            /fecha\s+de\s+creaci[oó]n/i,
            /creaci[oó]n\s+del?\s+perfil/i,
            /creaci[oó]n\s+de\s+la\s+cuenta/i,
            /creaci[oó]n\s+de\s+la\s+p[aá]gina/i,
            /perfil\s+creado/i,
            /cuenta\s+creada/i,
            /p[aá]gina\s+creada/i,
            /grupo\s+creado/i,
            /profile\s+created/i,
            /page\s+created/i,
            /account\s+created/i,
            /group\s+created/i,
            /se\s+uni[oó]\s+a\s+facebook/i,
            /se\s+uni[oó]\s+en/i,
            /se\s+uni[oó]/i,
            /joined\s+facebook/i,
            /joined\s+in/i,
            /joined/i,
            /fecha\s+de\s+incorporaci[oó]n/i,
            /member\s+since/i,
            /miembro\s+desde/i
        ];


        const labeledRaw = extractLabeledValue(body, labels);


        if (labeledRaw) {
            const formatted = formatMonthYearEs(labeledRaw);
            if (formatted !== 'N/D') return formatted;
        }


        return extractFacebookCreationDateFromSource();
    }


    function extractFacebookCreationDateFromSource() {
        const sources = getEmbeddedTextSources();


        for (const source of sources) {
            for (const variant of embeddedSourceVariants(source)) {
                const textMatches = variant.matchAll(/"text"\s*:\s*"([^"]{4,120})"/gi);
                for (const match of textMatches) {
                    const textVal = match[1];
                    if (
                        /se\s+uni[oó]|joined|fecha\s+de\s+creaci[oó]n|creaci[oó]n\s+del?\s+(?:perfil|cuenta|p[aá]gina|grupo)|(?:perfil|cuenta|p[aá]gina|grupo)\s+cread[oa]|created\s+on|profile\s+created|page\s+created|member\s+since|miembro\s+desde/i.test(textVal)
                    ) {
                        const parsed = extractFacebookCreationDateFromRawString(textVal);
                        if (parsed && parsed !== 'N/D') {
                            return parsed;
                        }
                    }
                }


                // Cuando en GraphQL un nodo tiene "text":"Fecha de creación del perfil" y el siguiente "text":"mayo de 2018"
                const adjacentTextPattern = /"text"\s*:\s*"(?:Fecha de creaci[oó]n[^"]*|Creaci[oó]n del perfil|Se uni[oó] a Facebook|Profile created|Page created|Date joined)"[\s\S]{0,350}?"text"\s*:\s*"([^"]{4,60})"/gi;
                let adjMatch;
                while ((adjMatch = adjacentTextPattern.exec(variant)) !== null) {
                    const parsed = formatMonthYearEs(adjMatch[1]);
                    if (parsed !== 'N/D') return parsed;
                }


                const labelPattern = /(?:fecha\s+de\s+creaci[oó]n|creaci[oó]n\s+del?\s+(?:perfil|cuenta|p[aá]gina|grupo)|profile\s+created|page\s+created|account\s+created)[^}]{0,300}?"(?:text|value|subtitle|title)"\s*:\s*"([^"]{4,60})"/gi;
                let labelMatch;
                while ((labelMatch = labelPattern.exec(variant)) !== null) {
                    const parsed = formatMonthYearEs(labelMatch[1]);
                    if (parsed !== 'N/D') return parsed;
                }


                const specificDatePattern = /"(?:page_creation_date|formatted_date_created|date_joined|join_date|creation_time_text)"\s*:\s*"?([^"}{]{4,60})"?/gi;
                let dateMatch;
                while ((dateMatch = specificDatePattern.exec(variant)) !== null) {
                    const parsed = formatMonthYearEs(dateMatch[1]);
                    if (parsed !== 'N/D') return parsed;
                }


                const specificEpochPattern = /"(?:profile_creation_time|account_creation_time|user_creation_time|registration_time)"\s*:\s*(\d{9,11})\b/gi;
                let epochMatch;
                while ((epochMatch = specificEpochPattern.exec(variant)) !== null) {
                    const parsed = formatEpochMonthYearEs(Number(epochMatch[1]));
                    if (parsed !== 'N/D') return parsed;
                }
            }
        }


        return 'N/D';
    }


    function extractFacebookCreationDateFromRawString(text) {
        if (!text) return 'N/D';


        const patterns = [
            /se\s+uni[oó]\s+en\s+([A-Za-zÁÉÍÓÚÜÑáéíóúüñ]+\s+(?:del?\s+)?\d{4})/i,
            /se\s+uni[oó]\s+en\s+(\d{4})/i,
            /se\s+uni[oó](?:\s+a\s+facebook)?\s+el\s+([^.\n.,;]{3,80})/i,
            /se\s+uni[oó]\s+a\s+facebook(?:\s+en)?\s+([^\n.,;]{3,80})/i,
            /joined(?:\s+facebook)?(?:\s+in|\s+on)?\s+([A-Za-z]+\s+\d{1,2},?\s+\d{4})/i,
            /joined(?:\s+facebook)?(?:\s+in|\s+on)?\s+([A-Za-z]+\s+\d{4})/i,
            /joined(?:\s+facebook)?(?:\s+in|\s+on)?\s+(\d{4})/i,
            /fecha\s+de\s+creaci[oó]n\s+del?\s+(?:perfil|cuenta|p[aá]gina|grupo)\s*:?\s*([^\n.,;]{3,80})/i,
            /fecha\s+de\s+creaci[oó]n\s*:?\s*([^\n.,;]{3,80})/i,
            /(?:perfil|p[aá]gina|cuenta|grupo)\s+cread[oa](?:\s+el|\s+en)?\s*:?\s*([^\n.,;]{3,80})/i,
            /(?:profile|page|account|group)\s+created(?:\s+on|\s+in)?\s*:?\s*([^\n.,;]{3,80})/i,
            /created\s+on\s+([^\n.,;]{3,80})/i,
            /member\s+since\s*:?\s*([^\n.,;]{3,80})/i,
            /miembro\s+desde\s*:?\s*([^\n.,;]{3,80})/i
        ];


        for (const pattern of patterns) {
            const match = text.match(pattern);
            if (match?.[1]) {
                const formatted = formatMonthYearEs(match[1]);
                if (formatted !== 'N/D') return formatted;
            }
        }


        return formatMonthYearEs(text);
    }


    function extractFacebookGroupName() {
        const heading = Array.from(
            document.querySelectorAll('h1')
        ).find(isVisible);


        if (heading) {
            const value = clean(
                heading.innerText || heading.textContent || ''
            );


            if (value) return value;
        }


        return cleanFacebookTitle(
            metaContent('meta[property="og:title"]') ||
            document.title
        );
    }


    function extractFacebookDisplayName() {
        const heading = Array.from(
            document.querySelectorAll('h1')
        ).find(isVisible);


        if (heading) {
            return clean(
                heading.innerText || heading.textContent || ''
            );
        }


        return cleanFacebookTitle(
            metaContent('meta[property="og:title"]') ||
            document.title
        );
    }


    function cleanFacebookTitle(value) {
        return clean(value)
            .replace(/\s*\|\s*Facebook\s*$/i, '')
            .replace(/\s*-\s*Facebook\s*$/i, '');
    }


    /********************************************************************
     * TIKTOK
     ********************************************************************/


    async function extractTikTok() {
        const username =
            getUserFromUrl() ||
            text('[data-e2e="user-title"]') ||
            'N/D';


        let bundle = findTikTokUserBundle(username);


        if (bundle.notFound || detectNotFound()) {
            return buildUnavailableRow(
                username,
                'NO',
                'El perfil de TikTok no existe, ha sido eliminado o está suspendido.'
            );
        }


        const creationKeys = [
            'createTime',
            'create_time',
            'createdAt',
            'created_at',
            'accountCreateTime'
        ];
        const nicknameChangeKeys = [
            'nickNameModifyTime',
            'nicknameModifyTime',
            'nickname_modify_time',
            'nick_name_modify_time'
        ];
        const pageUser = bundle.user || {};
        let fetchedProfileHtml = false;


        if (
            !pick(pageUser, ['id', 'uid', 'userId', 'user_id']) ||
            !pick(pageUser, creationKeys) ||
            !pick(pageUser, nicknameChangeKeys)
        ) {
            const fetchedBundle = await fetchTikTokUserBundle(username);


            if (fetchedBundle.notFound) {
                return buildUnavailableRow(
                    username,
                    'NO',
                    'El perfil de TikTok no existe (confirmado al consultar el estado de la cuenta).'
                );
            }


            if (fetchedBundle.user) {
                bundle = {
                    user: {
                        ...pageUser,
                        ...fetchedBundle.user
                    },
                    stats: bundle.stats || fetchedBundle.stats,
                    notFound: false
                };
                fetchedProfileHtml = true;
            }
        }


        const user = bundle.user || {};
        const stats = bundle.stats || {};


        const internalId =
            pick(user, ['id', 'uid', 'userId', 'user_id']) ||
            'N/D';


        const domUserTitle = text('[data-e2e="user-title"]');


        // Si tras analizar DOM y JSON no hay ID ni título de usuario en absoluto, la cuenta no existe
        if (internalId === 'N/D' && !domUserTitle) {
            return buildUnavailableRow(
                username,
                'NO',
                'No se encontraron datos de usuario ni elementos de perfil en TikTok.'
            );
        }


        const followers =
            pick(stats, [
                'followerCount',
                'follower_count',
                'followers'
            ]) ||
            text('[data-e2e="followers-count"]') ||
            extractMetric(getBodyText(), [
                /([\d][\d.,\s]*(?:K|M|B|mil|millones)?)\s*(?:Followers|Seguidores)\b/i
            ]);


        const following =
            pick(stats, [
                'followingCount',
                'following_count',
                'followings',
                'following'
            ]) ||
            text('[data-e2e="following-count"]') ||
            extractFollowingMetric(getBodyText());


        const creationEpoch = pick(user, creationKeys);
        const nicknameChangeEpoch = pick(user, nicknameChangeKeys);


        const creationDate =
            formatEpochMonthYearEs(creationEpoch);


        const nicknameChangeDate =
            formatEpochMonthYearEs(nicknameChangeEpoch);


        const location = await resolveTikTokLocation(username, user);


        const missingRestrictedFields =
            creationDate === 'N/D' ||
            nicknameChangeDate === 'N/D';


        return {
            USUARIO: username,
            'NOMBRE VISIBLE':
                pick(user, [
                    'nickname',
                    'nickName',
                    'displayName',
                    'display_name'
                ]) ||
                text('[data-e2e="user-subtitle"]') ||
                'N/D',
            'ID interno': internalId,
            'Nº SEGUIDORES': String(followers || 'N/D'),
            'Nº SEGUIDOS': String(following || 'N/D'),
            'TIPO DE OBJETO': 'PERFIL',
            MÉTRICA: 'SEGUIDORES',
            PLATAFORMA: PLATFORM,
            'FECHA DE CREACIÓN DE CUENTA': creationDate,
            'UBICACIÓN DE LA CUENTA': location,
            'ÚLTIMO CAMBIO DE USUARIO': nicknameChangeDate,
            'Nº CAMBIOS DE USUARIO': 'N/D',
            EXISTE: 'SI',
            URL: canonicalProfileUrl(),
            OBSERVACIONES: missingRestrictedFields
                ? 'TikTok no ha publicado en esta carga la fecha de creación o el cambio de nick.'
                : fetchedProfileHtml
                    ? 'Datos obtenidos del objeto exacto del perfil devuelto por TikTok.'
                    : 'Datos obtenidos del objeto exacto del usuario cargado en la página.'
        };
    }


    async function resolveTikTokLocation(username, user) {
        // 1. Propiedades directas del objeto user (si TikTok las incluye)
        const directLocation = pick(user, [
            'region',
            'accountRegion',
            'registerRegion',
            'country',
            'countryCode',
            'country_code',
            'location'
        ]);


        if (isUsefulValue(directLocation)) {
            return clean(directLocation);
        }


        // 2. Buscar en texto/diálogo si "Acerca de esta cuenta" está visible en el DOM
        const body = getBodyText();
        const visibleLocation = rxFirst(body, [
            /(?:Account based in|Cuenta basada en|Ubicación de la cuenta|País o región)\s*:?\s*([^\n]{2,60})/i
        ]) || extractLabeledValue(body, [
            /Account based in/i,
            /Cuenta basada en/i,
            /Ubicación de la cuenta/i
        ]);


        if (isUsefulValue(visibleLocation)) {
            return cleanAccountLocation(visibleLocation);
        }


        // 3. Buscar "locationCreated" en los vídeos cargados en los scripts de la página
        const embeddedVideoRegion = findTikTokVideoLocationInScripts(username);
        if (embeddedVideoRegion) {
            return embeddedVideoRegion;
        }


        // 4. Obtener un vídeo reciente desde el DOM o desde /embed/@username y leer su "locationCreated"
        const fetchedVideoRegion = await fetchTikTokRegionFromVideoOrEmbed(username);
        if (fetchedVideoRegion) {
            return fetchedVideoRegion;
        }


        return 'N/D';
    }


    function findTikTokVideoLocationInScripts(username) {
        const roots = [
            parseJsonScript('SIGI_STATE'),
            parseJsonScript('__UNIVERSAL_DATA_FOR_REHYDRATION__'),
            parseJsonScript('__NEXT_DATA__')
        ].filter(Boolean);


        let foundRegion = '';


        for (const root of roots) {
            walkObject(root, obj => {
                if (foundRegion || !obj || typeof obj !== 'object') return;


                const authorUsername = clean(
                    obj.author?.uniqueId ||
                    obj.author?.unique_id ||
                    (typeof obj.author === 'string' ? obj.author : '') ||
                    ''
                );


                if (
                    authorUsername &&
                    sameIdentity(authorUsername, username)
                ) {
                    const candidate = clean(
                        obj.locationCreated ||
                        obj.location_created ||
                        obj.poi?.countryCode ||
                        obj.poi?.country ||
                        ''
                    );


                    if (candidate && /^[A-Z]{2,3}$/i.test(candidate)) {
                        foundRegion = candidate.toUpperCase();
                    }
                }
            });
        }


        return foundRegion;
    }


    async function fetchTikTokRegionFromVideoOrEmbed(username) {
        if (!username || username === 'N/D') return '';


        try {
            // Intentar localizar el primer enlace de vídeo del propio usuario en el DOM
            const escapedUser = escapeRegExp(username);
            const videoLinkPattern = new RegExp(`/@${escapedUser}/video/(\\d+)`, 'i');


            let videoUrl = '';
            const domLinks = Array.from(document.querySelectorAll('a[href*="/video/"]'));


            for (const link of domLinks) {
                const href = link.getAttribute('href') || '';
                if (videoLinkPattern.test(href)) {
                    videoUrl = new URL(href, location.origin).href;
                    break;
                }
            }


            // Si el DOM aún no tiene el enlace del vídeo, consultar la página ligera /embed/@username
            if (!videoUrl) {
                const embedResp = await fetch(
                    `https://www.tiktok.com/embed/@${encodeURIComponent(username)}`,
                    { credentials: 'include' }
                );


                if (embedResp.ok) {
                    const embedHtml = await embedResp.text();


                    const directLoc = embedHtml.match(/"locationCreated"\s*:\s*"([A-Za-z]{2,3})"/i);
                    if (directLoc?.[1]) {
                        return directLoc[1].toUpperCase();
                    }


                    const videoIdMatch =
                        embedHtml.match(new RegExp(`/@${escapedUser}/video/(\\d{15,})`, 'i')) ||
                        embedHtml.match(/"video_id"\s*:\s*"(\d{15,})"/i) ||
                        embedHtml.match(/\/video\/(\d{15,})/i);


                    if (videoIdMatch?.[1]) {
                        videoUrl = `https://www.tiktok.com/@${encodeURIComponent(username)}/video/${videoIdMatch[1]}`;
                    }
                }
            }


            if (!videoUrl) return '';


            const videoResp = await fetch(videoUrl, {
                credentials: 'include',
                headers: { accept: 'text/html,application/xhtml+xml' }
            });


            if (!videoResp.ok) return '';


            const videoHtml = await videoResp.text();
            const locMatch =
                videoHtml.match(/"locationCreated"\s*:\s*"([A-Za-z]{2,3})"/i) ||
                videoHtml.match(/"location_created"\s*:\s*"([A-Za-z]{2,3})"/i);


            return locMatch?.[1] ? locMatch[1].toUpperCase() : '';
        } catch {
            return '';
        }
    }


    async function fetchTikTokUserBundle(username) {
        if (!username || username === 'N/D') {
            return { user: null, stats: null, notFound: false };
        }


        try {
            const response = await fetch(
                `https://www.tiktok.com/@${encodeURIComponent(username)}`,
                {
                    credentials: 'include',
                    cache: 'no-store',
                    headers: {
                        accept: 'text/html,application/xhtml+xml'
                    }
                }
            );


            if (response.status === 404) {
                return { user: null, stats: null, notFound: true };
            }


            if (!response.ok) {
                throw new Error(`HTTP ${response.status}`);
            }


            return findTikTokUserBundleInHtml(
                await response.text(),
                username
            );
        } catch (error) {
            console.warn(
                '[Roomba Scraper] No se pudo releer el perfil de TikTok:',
                error
            );


            return { user: null, stats: null, notFound: false };
        }
    }


    function findTikTokUserBundleInHtml(html, username) {
        const parsedDocument = new DOMParser().parseFromString(
            html,
            'text/html'
        );
        const roots = [
            'SIGI_STATE',
            '__UNIVERSAL_DATA_FOR_REHYDRATION__',
            '__NEXT_DATA__'
        ].map(id => {
            const content = parsedDocument.getElementById(id)?.textContent;


            if (!content) return null;


            try {
                return JSON.parse(content);
            } catch {
                return null;
            }
        }).filter(Boolean);
        const structured = findTikTokUserBundle(username, roots);
        const rawUser = findTikTokUserObjectInHtml(html, username);


        return {
            user: structured.user || rawUser
                ? {
                    ...(structured.user || {}),
                    ...(rawUser || {})
                }
                : null,
            stats: structured.stats,
            notFound: structured.notFound && !rawUser
        };
    }


    function findTikTokUserObjectInHtml(html, username) {
        const usernameJson = JSON.stringify(String(username));
        const marker = new RegExp(
            `"uniqueId"\\s*:\\s*${escapeRegExp(usernameJson)}`,
            'g'
        );
        let match;


        while ((match = marker.exec(html))) {
            const scriptStart = html.lastIndexOf('<script', match.index);
            const contentStart = scriptStart >= 0
                ? html.indexOf('>', scriptStart) + 1
                : -1;
            const contentEnd = contentStart > 0
                ? html.indexOf('</script>', match.index)
                : -1;


            if (contentStart <= 0 || contentEnd <= match.index) {
                continue;
            }


            const scriptText = html.slice(contentStart, contentEnd);
            const markerIndex = match.index - contentStart;
            const openings = findOpenJsonObjects(scriptText, markerIndex);


            for (let index = openings.length - 1; index >= 0; index -= 1) {
                const object = parseJsonObjectAt(
                    scriptText,
                    openings[index]
                );
                const objectUsername = clean(
                    object?.uniqueId ||
                    object?.unique_id ||
                    object?.username ||
                    ''
                );


                if (
                    objectUsername &&
                    objectUsername.toLowerCase() === username.toLowerCase()
                ) {
                    return object;
                }
            }
        }


        return null;
    }


    function findOpenJsonObjects(source, endIndex) {
        const stack = [];
        let inString = false;
        let escaped = false;


        for (let index = 0; index < endIndex; index += 1) {
            const character = source[index];


            if (inString) {
                if (escaped) {
                    escaped = false;
                } else if (character === '\\') {
                    escaped = true;
                } else if (character === '"') {
                    inString = false;
                }


                continue;
            }


            if (character === '"') {
                inString = true;
            } else if (character === '{') {
                stack.push(index);
            } else if (character === '}') {
                stack.pop();
            }
        }


        return stack;
    }


    function parseJsonObjectAt(source, startIndex) {
        let depth = 0;
        let inString = false;
        let escaped = false;


        for (let index = startIndex; index < source.length; index += 1) {
            const character = source[index];


            if (inString) {
                if (escaped) {
                    escaped = false;
                } else if (character === '\\') {
                    escaped = true;
                } else if (character === '"') {
                    inString = false;
                }


                continue;
            }


            if (character === '"') {
                inString = true;
            } else if (character === '{') {
                depth += 1;
            } else if (character === '}') {
                depth -= 1;


                if (depth === 0) {
                    try {
                        return JSON.parse(source.slice(startIndex, index + 1));
                    } catch {
                        return null;
                    }
                }
            }
        }


        return null;
    }


    function findTikTokUserBundle(username, suppliedRoots = null) {
        const roots = suppliedRoots || [
            parseJsonScript('SIGI_STATE'),
            parseJsonScript('__UNIVERSAL_DATA_FOR_REHYDRATION__'),
            parseJsonScript('__NEXT_DATA__')
        ].filter(Boolean);


        let matchedUser = null;
        let matchedStats = null;
        let explicitNotFound = false;


        for (const root of roots) {
            // Verificar statusCode oficial en __UNIVERSAL_DATA_FOR_REHYDRATION__
            const userDetail = root?.__DEFAULT_SCOPE__?.['webapp.user-detail'];
            if (userDetail && typeof userDetail === 'object') {
                const statusCode = Number(userDetail.statusCode ?? 0);
                const hasUserObj = Boolean(
                    userDetail.userInfo?.user?.id ||
                    userDetail.userInfo?.user?.uniqueId
                );


                // Códigos de TikTok cuando una cuenta no existe, está borrada o baneada (ej. 10202, 10221, 10222, 10223)
                if (statusCode !== 0 && !hasUserObj) {
                    explicitNotFound = true;
                }
            }


            walkObject(root, object => {
                if (!object || typeof object !== 'object') return;


                if (object.user && typeof object.user === 'object') {
                    const nestedUsername = clean(
                        object.user.uniqueId ||
                        object.user.unique_id ||
                        object.user.username ||
                        ''
                    );


                    if (
                        nestedUsername &&
                        nestedUsername.toLowerCase() === username.toLowerCase()
                    ) {
                        matchedUser ||= object.user;


                        if (object.stats && typeof object.stats === 'object') {
                            matchedStats ||= object.stats;
                        }
                    }
                }


                const directUsername = clean(
                    object.uniqueId ||
                    object.unique_id ||
                    object.username ||
                    ''
                );


                if (
                    directUsername &&
                    directUsername.toLowerCase() === username.toLowerCase()
                ) {
                    matchedUser ||= object;
                }
            });
        }


        return {
            user: matchedUser,
            stats: matchedStats,
            notFound: explicitNotFound && !matchedUser
        };
    }


    /********************************************************************
     * THREADS
     ********************************************************************/


    function extractThreads() {
        const username = getUserFromUrl() || 'N/D';


        const description =
            metaContent('meta[property="og:description"]') ||
            metaContent('meta[name="description"]') ||
            '';


        const followers =
            rxFirst(description, [
                /([\d][\d.,\s]*(?:K|M|B|mil|millones)?)\s+(?:followers|seguidores)/i
            ]) ||
            extractMetric(getBodyText(), [
                /([\d][\d.,\s]*(?:K|M|B|mil|millones)?)\s+(?:followers|seguidores)/i
            ]);


        const following =
            extractFollowingMetric(description) ||
            extractFollowingMetric(getBodyText());


        return {
            USUARIO: username,
            'ID interno': findThreadsInternalId(username),
            'Nº SEGUIDORES': followers || 'N/D',
            'Nº SEGUIDOS': following || 'N/D',
            'TIPO DE OBJETO': 'PERFIL',
            MÉTRICA: 'SEGUIDORES',
            PLATAFORMA: PLATFORM,
            'FECHA DE CREACIÓN DE CUENTA': 'N/D',
            'UBICACIÓN DE LA CUENTA': 'N/D',
            'ÚLTIMO CAMBIO DE USUARIO': 'N/D',
            'Nº CAMBIOS DE USUARIO': 'N/D',
            EXISTE: 'SI',
            URL: canonicalProfileUrl(),
            OBSERVACIONES:
                'Threads no expone de forma estable todos los campos solicitados en el perfil web.'
        };
    }


    /********************************************************************
     * REANUDACIÓN DESPUÉS DE NAVEGAR A /ABOUT
     ********************************************************************/


    async function resumePendingCapture() {
        const pending = getPending();


        if (!pending || pending.platform !== PLATFORM) {
            return false;
        }


        if (
            pending.stage === 'facebook-id-fallback' &&
            PLATFORM === 'Facebook' &&
            /^\/groups\//i.test(location.pathname) &&
            sameIdentity(
                pending.identity,
                location.pathname.match(/^\/groups\/([^/?#]+)/i)?.[1] || ''
            )
        ) {
            await waitForPage();


            const fromQueue = pending.fromQueue;
            const identity = pending.identity;
            clearPending();


            if (detectBlock()) {
                pauseQueue();
                setStatus(
                    'Captura pausada por bloqueo o verificación en Facebook.'
                );
                return true;
            }


            if (detectNotFound()) {
                await commitRow(
                    buildUnavailableRow(
                        identity,
                        'NO',
                        'El ID no corresponde a un perfil ni a un grupo accesible de Facebook.'
                    ),
                    fromQueue
                );
                return true;
            }


            await captureCurrentProfile(fromQueue);
            return true;
        }


        if (
            pending.stage === 'x-about' &&
            PLATFORM === 'Twitter/X' &&
            isXAboutPage() &&
            sameIdentity(pending.identity, getUserFromUrl())
        ) {
            await waitForPage();


            if (detectBlock()) {
                pauseQueue();
                clearPending();
                setStatus('Captura pausada por bloqueo o verificación en X.');
                return true;
            }


            const mergedAboutData = mergeXAboutData(
                pending.base,
                await extractXAboutPageWithHistory()
            );


            clearPending();
            await commitRow(mergedAboutData, pending.fromQueue);
            return true;
        }


        if (
            pending.stage === 'facebook-about' &&
            PLATFORM === 'Facebook' &&
            isFacebookAboutPage() &&
            (
                sameIdentity(pending.identity, getFacebookIdentity()) ||
                sameIdentity(pending.base?.['ID interno'], getFacebookIdentity())
            )
        ) {
            await waitForPage();


            if (detectBlock()) {
                pauseQueue();
                clearPending();
                setStatus(
                    'Captura pausada por bloqueo o verificación en Facebook.'
                );
                return true;
            }


            const finalData = mergeRows(
                pending.base,
                extractFacebookPage()
            );


            clearPending();
            await commitRow(finalData, pending.fromQueue);
            return true;
        }


        return false;
    }


    function savePending(data) {
        const pending = {
            ...data,
            createdAt: Date.now()
        };


        GM_setValue(PENDING_KEY, JSON.stringify(pending));
    }


    function getPending() {
        try {
            const raw = GM_getValue(PENDING_KEY, '');


            if (!raw) return null;


            const pending = JSON.parse(raw);


            if (
                !pending.createdAt ||
                Date.now() - pending.createdAt > 15 * 60 * 1000
            ) {
                clearPending();
                return null;
            }


            return pending;
        } catch {
            clearPending();
            return null;
        }
    }


    function clearPending() {
        GM_deleteValue(PENDING_KEY);
    }


    /********************************************************************
     * COLA
     ********************************************************************/


    async function runAutoQueueSafely() {
        try {
            await runAutoQueue();
        } catch (error) {
            console.error('[Roomba Scraper] Error en cola:', error);
            pauseQueue();


            setStatus(
                `Cola pausada por error:\n${error.message || error}`
            );
        }
    }


    async function runAutoQueue() {
        const queue = getQueue();


        if (!queue.active || !queue.items.length) return;


        if (getPending()) {
            return;
        }


        if (detectBlock()) {
            pauseQueue();
            setStatus(
                'Cola pausada: se ha detectado bloqueo, límite o verificación.'
            );
            return;
        }


        if (queue.index >= queue.items.length) {
            queue.active = false;
            saveQueue(queue);
            setStatus('Cola finalizada.');
            updatePanelCounts();
            return;
        }


        const item = queue.items[queue.index];


        if (!item?.url) {
            queue.index += 1;
            saveQueue(queue);
            return runAutoQueue();
        }


        if (!sameProfileTarget(item.url, location.href)) {
            setStatus(
                `Abriendo perfil ${queue.index + 1}/${queue.items.length}:\n` +
                `${item.label}`
            );


            setTimeout(() => {
                location.href = item.url;
            }, 1300);


            return;
        }


        setStatus(
            `Capturando ${queue.index + 1}/${queue.items.length}:\n` +
            `${item.label}`
        );


        await sleep(randomInt(2500, 4500));
        await captureCurrentProfile(true);
    }


    async function advanceQueueAfterCommit(row) {
        const queue = getQueue();


        if (!queue.active || !queue.items.length) return;


        const currentItem = queue.items[queue.index];


        if (currentItem) {
            currentItem.status = row.EXISTE;
            currentItem.lastCapture = new Date().toISOString();
        }


        queue.index += 1;


        if (queue.index >= queue.items.length) {
            queue.active = false;
            saveQueue(queue);


            setStatus(
                `Cola finalizada.\nFilas guardadas: ${getRows().length}`
            );


            updatePanelCounts();
            return;
        }


        saveQueue(queue);
        refreshQueueTextarea();
        updatePanelCounts();


        const seconds = randomInt(
            AUTO_MIN_SECONDS,
            AUTO_MAX_SECONDS
        );


        setStatus(
            `Perfil registrado.\n` +
            `Siguiente perfil en ${seconds} segundos.\n` +
            `Pendientes: ${queue.items.length - queue.index}`
        );


        setTimeout(() => {
            const updatedQueue = getQueue();


            if (
                updatedQueue.active &&
                updatedQueue.items[updatedQueue.index]
            ) {
                location.href =
                    updatedQueue.items[updatedQueue.index].url;
            }
        }, seconds * 1000);
    }


    function saveQueueFromTextarea(showConfirmation = true) {
        const raw = $('#falconai-queue-input')?.value || '';


        const lines = raw
            .split(/\r?\n/)
            .map(clean)
            .filter(Boolean);


        const items = [];


        for (const line of lines) {
            const url = toProfileUrl(line);


            if (!url) continue;


            items.push({
                label: line,
                url,
                status: 'PENDIENTE',
                lastCapture: ''
            });
        }


        saveQueue({
            active: false,
            index: 0,
            createdAt: new Date().toISOString(),
            items
        });


        if (showConfirmation) {
            setStatus(`Cola guardada: ${items.length} perfiles.`);
        }


        updatePanelCounts();
    }


    function getQueue() {
        try {
            const raw = GM_getValue(QUEUE_KEY, '');


            if (!raw) {
                return {
                    active: false,
                    index: 0,
                    items: []
                };
            }


            const queue = JSON.parse(raw);


            return {
                active: Boolean(queue.active),
                index: Number(queue.index || 0),
                items: Array.isArray(queue.items)
                    ? queue.items
                    : []
            };
        } catch {
            return {
                active: false,
                index: 0,
                items: []
            };
        }
    }


    function saveQueue(queue) {
        GM_setValue(QUEUE_KEY, JSON.stringify(queue));
    }


    function pauseQueue() {
        const queue = getQueue();
        queue.active = false;
        saveQueue(queue);
        updatePanelCounts();
    }


    function refreshQueueTextarea() {
        const element = $('#falconai-queue-input');
        if (!element) return;


        const queue = getQueue();


        element.value = queue.items
            .map(item => item.label || item.url)
            .join('\n');
    }


    /********************************************************************
     * ALMACENAMIENTO Y EXCEL
     ********************************************************************/


    function getRows() {
        try {
            const raw = GM_getValue(DATA_KEY, '[]');
            const parsed = JSON.parse(raw);


            return Array.isArray(parsed) ? parsed : [];
        } catch {
            return [];
        }
    }


    function saveRows(rows) {
        GM_setValue(DATA_KEY, JSON.stringify(rows));
    }


    function upsertOrAppend(row, replaceFields = []) {
        const rows = getRows();
        let action = 'Añadido';


        if (RECORD_MODE === 'upsert') {
            const key = rowKey(row);
            const normalizedUser = normalizeText(row.USUARIO);
            const normalizedType = normalizeText(row['TIPO DE OBJETO']);


            const index = rows.findIndex(existing => {
                if (rowKey(existing) === key) return true;


                // Evitar duplicados si antes tenía ID = N/D y ahora tiene ID numérico (o viceversa)
                if (
                    isUsefulValue(normalizedUser) &&
                    normalizeText(existing.PLATAFORMA || PLATFORM) === normalizeText(row.PLATAFORMA || PLATFORM) &&
                    normalizeText(existing['TIPO DE OBJETO']) === normalizedType &&
                    normalizeText(existing.USUARIO) === normalizedUser
                ) {
                    return true;
                }


                return false;
            });


            if (index >= 0) {
                rows[index] = mergeRows(rows[index], row);


                for (const field of replaceFields) {
                    rows[index][field] = row[field];
                }


                action = 'Actualizado';
            } else {
                rows.push(row);
            }
        } else {
            rows.push(row);
        }


        saveRows(rows);


        return {
            action,
            total: rows.length
        };
    }


    function findExistingRow(username, type, internalId = '') {
        const normalizedUsername = normalizeText(username);
        const normalizedType = normalizeText(type);
        const cleanedId = clean(internalId);


        return getRows().find(row => {
            if (normalizeText(row['TIPO DE OBJETO']) !== normalizedType) {
                return false;
            }


            if (isUsefulValue(cleanedId) && clean(row['ID interno']) === cleanedId) {
                return true;
            }


            return normalizeText(row.USUARIO) === normalizedUsername;
        }) || null;
    }


    function exportHeaders() {
        return getPlatformHeaders().map(exportHeaderLabel);
    }


    function exportHeaderLabel(header) {
        return PLATFORM_HEADER_LABELS[PLATFORM]?.[header] || header;
    }


    function exportCellValue(row, header) {
        if (header === 'OBSERVACIONES') return '';


        if (header === 'URL') {
            return isUsefulValue(row.URL)
                ? String(row.URL)
                : 'N/D';
        }


        return String(row[header] ?? '');
    }


    function rawField(raw, header, aliases = []) {
        const candidates = [
            header,
            exportHeaderLabel(header),
            ...aliases
        ];


        for (const candidate of candidates) {
            if (
                raw[candidate] !== undefined &&
                raw[candidate] !== null &&
                raw[candidate] !== ''
            ) {
                return raw[candidate];
            }
        }


        return '';
    }


    function exportXlsx() {
        const rows = getRows();


        if (!rows.length) {
            setStatus('No hay filas para exportar.');
            return;
        }


        const XLSXLib =
            typeof XLSX !== 'undefined'
                ? XLSX
                : window.XLSX;


        if (!XLSXLib) {
            setStatus('No se ha cargado la librería XLSX.');
            return;
        }


        const sortedRows = [...rows];
        const activeHeaders = getPlatformHeaders();
        const visibleHeaders = activeHeaders.map(exportHeaderLabel);


        const matrix = [
            visibleHeaders,
            ...sortedRows.map(row => {
                return activeHeaders.map(header => {
                    return exportCellValue(row, header);
                });
            })
        ];


        const worksheet =
            XLSXLib.utils.aoa_to_sheet(matrix);


        sortedRows.forEach((row, index) => {
            const address = XLSXLib.utils.encode_cell({
                c: 0,
                r: index + 1
            });
            const target = clean(row.URL);


            if (worksheet[address] && target && !/^N\/D$/i.test(target)) {
                worksheet[address].l = {
                    Target: target,
                    Tooltip: target
                };
                worksheet[address].s = {
                    font: {
                        color: { rgb: '0563C1' },
                        underline: true
                    }
                };
            }
        });


        sortedRows.forEach((row, index) => {
            if (normalizeText(row.EXISTE) !== 'no') return;


            for (let columnIndex = 0; columnIndex < activeHeaders.length; columnIndex += 1) {
                const address = XLSXLib.utils.encode_cell({
                    c: columnIndex,
                    r: index + 1
                });


                if (!worksheet[address]) continue;


                worksheet[address].s = {
                    ...(worksheet[address].s || {}),
                    fill: {
                        patternType: 'solid',
                        fgColor: { rgb: 'FCE8E6' }
                    }
                };
            }
        });


        for (const address of Object.keys(worksheet)) {
            if (address.startsWith('!')) continue;


            worksheet[address].t = 's';
            worksheet[address].z = '@';
        }


        worksheet['!cols'] = activeHeaders.map(header => ({
            wch: COLUMN_WIDTHS[header] || 24
        }));


        const workbook = XLSXLib.utils.book_new();


        XLSXLib.utils.book_append_sheet(
            workbook,
            worksheet,
            'perfiles'
        );


        const xlsxName = buildXlsxName();


        XLSXLib.writeFile(workbook, xlsxName, {
            bookType: 'xlsx',
            cellStyles: true
        });


        setStatus(
            `Excel exportado:\n${xlsxName}\nFilas: ${rows.length}`
        );
    }


    async function importXlsx(file) {
        try {
            const XLSXLib =
                typeof XLSX !== 'undefined'
                    ? XLSX
                    : window.XLSX;


            if (!XLSXLib) {
                throw new Error(
                    'No se ha cargado la librería XLSX.'
                );
            }


            const buffer = await file.arrayBuffer();


            const workbook = XLSXLib.read(buffer, {
                type: 'array',
                cellDates: false
            });


            const worksheet =
                workbook.Sheets[workbook.SheetNames[0]];


            const imported = XLSXLib.utils.sheet_to_json(
                worksheet,
                {
                    defval: '',
                    raw: false
                }
            );


            const rows = getRows();
            let added = 0;
            let updated = 0;


            for (const importedRow of imported) {
                const normalized = normalizeRow(importedRow);


                if (normalized.PLATAFORMA !== PLATFORM) {
                    continue;
                }


                const key = rowKey(normalized);
                const index = rows.findIndex(
                    row => rowKey(row) === key
                );


                if (index >= 0) {
                    rows[index] = mergeRows(
                        rows[index],
                        normalized
                    );


                    updated += 1;
                } else {
                    rows.push(normalized);
                    added += 1;
                }
            }


            saveRows(rows);


            setStatus(
                `Importación completada.\n` +
                `Añadidas: ${added}\n` +
                `Actualizadas: ${updated}\n` +
                `Total: ${rows.length}`
            );


            updatePanelCounts();
        } catch (error) {
            console.error('[Roomba Scraper] Error importando:', error);


            setStatus(
                `Error importando el Excel:\n` +
                `${error.message || error}`
            );
        }
    }


    /********************************************************************
     * NORMALIZACIÓN
     ********************************************************************/


    function normalizeRow(raw = {}) {
        const row = {};


        for (const header of HEADERS) {
            row[header] = '';
        }


        row.USUARIO = clean(
            rawField(raw, 'USUARIO') ||
            raw.usuario ||
            getUserFromUrl() ||
            'N/D'
        ).replace(/^@/, '');


        row['NOMBRE VISIBLE'] = clean(
            rawField(raw, 'NOMBRE VISIBLE', [
                'FB: NOMBRE VISIBLE'
            ]) ||
            raw.displayName ||
            raw.nombreVisible ||
            'N/D'
        );


        row['ID interno'] = clean(
            rawField(raw, 'ID interno', [
                'FB: ID interno',
                'X: ID interno',
                'IG: ID interno',
                'TT: ID interno',
                'TH: ID interno'
            ]) ||
            raw.idInterno ||
            raw.id ||
            'N/D'
        );


        row['Nº SEGUIDORES'] = formatAudienceCount(
            rawField(raw, 'Nº SEGUIDORES', [
                'FB: AMIGOS/MIEMBROS'
            ]) ||
            raw.seguidores ||
            raw.amigos ||
            raw.miembros ||
            'N/D'
        );


        row['Nº SEGUIDOS'] = formatAudienceCount(
            rawField(raw, 'Nº SEGUIDOS', [
                'FB: SEGUIDOS',
                'X: SEGUIDOS',
                'IG: SEGUIDOS',
                'TT: SEGUIDOS',
                'TH: SEGUIDOS'
            ]) ||
            raw.seguidos ||
            raw.siguiendo ||
            raw.following ||
            raw.followingCount ||
            'N/D'
        );


        row['TIPO DE OBJETO'] = clean(
            rawField(raw, 'TIPO DE OBJETO') ||
            raw.tipo ||
            'PERFIL'
        ).toUpperCase();


        row.MÉTRICA = clean(
            rawField(raw, 'MÉTRICA') ||
            raw.metrica ||
            'SEGUIDORES'
        ).toUpperCase();


        row.PLATAFORMA = PLATFORM;


        row['FECHA DE CREACIÓN DE CUENTA'] =
            normalizeDateField(
                rawField(raw, 'FECHA DE CREACIÓN DE CUENTA', [
                    'FB: FECHA DE CREACIÓN DE CUENTA',
                    'X: FECHA DE CREACIÓN DE CUENTA',
                    'IG: FECHA DE CREACIÓN DE CUENTA',
                    'TT: FECHA DE CREACIÓN DE CUENTA'
                ]) ||
                raw.fechaCreacion ||
                'N/D'
            );


        row['UBICACIÓN DE LA CUENTA'] = clean(
            rawField(raw, 'UBICACIÓN DE LA CUENTA', [
                'FB: UBICACIÓN DE LA CUENTA',
                'X: UBICACIÓN DE LA CUENTA',
                'IG: UBICACIÓN DE LA CUENTA',
                'TT: UBICACIÓN DE LA CUENTA'
            ]) ||
            raw.ubicacion ||
            'N/D'
        );


        row['ÚLTIMO CAMBIO DE USUARIO'] =
            normalizeDateField(
                rawField(raw, 'ÚLTIMO CAMBIO DE USUARIO', [
                    'X: ÚLTIMO CAMBIO DE USUARIO',
                    'TT: ÚLTIMO CAMBIO DE USUARIO',
                    'TT: ÚLTIMO CAMBIO DE NICKNAME'
                ]) ||
                raw.ultimoCambioUsuario ||
                'N/D'
            );


        row['ÚLTIMA ACTUALIZACIÓN DEL PERFIL'] = clean(
            rawField(raw, 'ÚLTIMA ACTUALIZACIÓN DEL PERFIL', [
                'FB: ÚLTIMA ACTUALIZACIÓN DEL PERFIL'
            ]) ||
            raw.ultimaActualizacionPerfil ||
            'N/D'
        );


        row['Nº CAMBIOS DE USUARIO'] = clean(
            rawField(raw, 'Nº CAMBIOS DE USUARIO', [
                'X: Nº CAMBIOS DE USUARIO',
                'IG: Nº CAMBIOS DE USUARIO'
            ]) ||
            raw.numeroCambiosUsuario ||
            'N/D'
        );


        const existence = clean(
            rawField(raw, 'EXISTE') ||
            raw.existe ||
            'SI'
        ).toUpperCase();


        row.EXISTE = existence || 'SI';


        row.URL = clean(
            rawField(raw, 'URL') ||
            canonicalProfileUrl()
        );


        row['Motivación del perfil'] = clean(
            rawField(raw, 'Motivación del perfil') ||
            raw.motivacionPerfil ||
            ''
        );


        row.OBSERVACIONES = '';


        row['GRUPO PÚBLICO'] = clean(
            rawField(raw, 'GRUPO PÚBLICO', [
                'FB: GRUPO PÚBLICO'
            ]) ||
            raw.grupoPublico ||
            'N/D'
        ).toUpperCase();


        row['GRUPO VISIBLE'] = clean(
            rawField(raw, 'GRUPO VISIBLE', [
                'FB: GRUPO VISIBLE'
            ]) ||
            raw.grupoVisible ||
            'N/D'
        ).toUpperCase();


        for (const header of HEADERS) {
            if (
                header === 'Motivación del perfil' ||
                header === 'OBSERVACIONES'
            ) {
                row[header] = String(row[header] || '');
                continue;
            }


            row[header] = String(
                row[header] === '' ? 'N/D' : row[header]
            );
        }


        return row;
    }


    function normalizeDateField(value) {
        const cleaned = clean(value);


        if (cleaned === '0' || cleaned === '-') {
            return cleaned;
        }


        if (!cleaned || /^N\/D$/i.test(cleaned)) {
            return 'N/D';
        }


        return formatMonthYearEs(cleaned);
    }


    function formatMonthYearEs(raw) {
        const value = clean(raw);


        if (!value || /^N\/D$/i.test(value)) {
            return 'N/D';
        }


        const normalized = normalizeText(value);
        const militaryMatch = normalized.match(
            /^(0[1-9]|[12]\d|3[01])(ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)((?:19|20|21)\d{2})$/i
        );


        if (militaryMatch) {
            return (
                militaryMatch[1] +
                militaryMatch[2].toUpperCase() +
                militaryMatch[3]
            );
        }


        const militaryMonthMatch = normalized.match(
            /^(ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)((?:19|20|21)\d{2})$/i
        );


        if (militaryMonthMatch) {
            return (
                militaryMonthMatch[1].toUpperCase() +
                militaryMonthMatch[2]
            );
        }


        const yearMatch = normalized.match(
            /\b(19\d{2}|20\d{2}|21\d{2})\b/
        );


        if (!yearMatch) {
            return 'N/D';
        }


        const monthDefinitions = [
            {
                names: ['enero', 'ene', 'january', 'jan'],
                number: '01',
                output: 'Enero'
            },
            {
                names: ['febrero', 'feb', 'february'],
                number: '02',
                output: 'Febrero'
            },
            {
                names: ['marzo', 'mar', 'march'],
                number: '03',
                output: 'Marzo'
            },
            {
                names: ['abril', 'abr', 'april', 'apr'],
                number: '04',
                output: 'Abril'
            },
            {
                names: ['mayo', 'may'],
                number: '05',
                output: 'Mayo'
            },
            {
                names: ['junio', 'jun', 'june'],
                number: '06',
                output: 'Junio'
            },
            {
                names: ['julio', 'jul', 'july'],
                number: '07',
                output: 'Julio'
            },
            {
                names: ['agosto', 'ago', 'august', 'aug'],
                number: '08',
                output: 'Agosto'
            },
            {
                names: [
                    'septiembre',
                    'setiembre',
                    'sept',
                    'sep',
                    'september'
                ],
                number: '09',
                output: 'Septiembre'
            },
            {
                names: ['octubre', 'oct', 'october'],
                number: '10',
                output: 'Octubre'
            },
            {
                names: ['noviembre', 'nov', 'november'],
                number: '11',
                output: 'Noviembre'
            },
            {
                names: ['diciembre', 'dic', 'december', 'dec'],
                number: '12',
                output: 'Diciembre'
            }
        ];


        for (const month of monthDefinitions) {
            const monthPattern = month.names
                .map(escapeRegExp)
                .join('|');
            const dayBeforeMonth = normalized.match(
                new RegExp(
                    `\\b(0?[1-9]|[12]\\d|3[01])\\s*(?:de\\s+)?(?:${monthPattern})(?:\\s+de)?\\s*,?\\s*${yearMatch[1]}\\b`,
                    'i'
                )
            );
            const dayAfterMonth = normalized.match(
                new RegExp(
                    `\\b(?:${monthPattern})\\s+(0?[1-9]|[12]\\d|3[01])(?:st|nd|rd|th)?\\s*,?\\s*${yearMatch[1]}\\b`,
                    'i'
                )
            );


            if (
                month.names.some(name => {
                    const pattern = new RegExp(
                        `(^|[^a-z])${escapeRegExp(name)}([^a-z]|$)`,
                        'i'
                    );


                    return pattern.test(normalized);
                })
            ) {
                const day = dayBeforeMonth?.[1] || dayAfterMonth?.[1] || '';


                if (day) {
                    return formatMilitaryDate(day, month.number, yearMatch[1]);
                }


                return formatMilitaryDate(
                    '01',
                    month.number,
                    yearMatch[1]
                ).slice(2);
            }
        }


        const isoMatch = normalized.match(
            /\b((?:19|20|21)\d{2})[-/](0[1-9]|1[0-2])(?:[-/](0?[1-9]|[12]\d|3[01]))?\b/
        );


        if (isoMatch) {
            return isoMatch[3]
                ? formatMilitaryDate(isoMatch[3], isoMatch[2], isoMatch[1])
                : formatMilitaryDate(
                    '01',
                    isoMatch[2],
                    isoMatch[1]
                ).slice(2);
        }


        const europeanMatch = normalized.match(
            /\b(0?[1-9]|[12]\d|3[01])[-/](0?[1-9]|1[0-2])[-/]((?:19|20|21)\d{2})\b/
        );


        if (europeanMatch) {
            return formatMilitaryDate(
                europeanMatch[1],
                europeanMatch[2],
                europeanMatch[3]
            );
        }


        // Si Facebook u otra red ofrece únicamente el año (ej. "Se unió en 2016"), devolver el año
        if (/^(?:en\s+|in\s+|del?\s+)?(19\d{2}|20\d{2}|21\d{2})$/i.test(normalized)) {
            return yearMatch[1];
        }


        return 'N/D';
    }


    function formatEpochMonthYearEs(value) {
        if (value === undefined || value === null || value === '') {
            return 'N/D';
        }


        const numeric = Number(value);


        if (!Number.isFinite(numeric) || numeric <= 0) {
            return 'N/D';
        }


        const milliseconds =
            numeric < 100000000000
                ? numeric * 1000
                : numeric;


        const date = new Date(milliseconds);


        if (
            Number.isNaN(date.getTime()) ||
            date.getUTCFullYear() < 2000 ||
            date.getUTCFullYear() > 2100
        ) {
            return 'N/D';
        }


        return formatMilitaryDate(
            date.getUTCDate(),
            date.getUTCMonth() + 1,
            date.getUTCFullYear()
        );
    }


    function formatAudienceCount(value) {
        const original = clean(value);


        if (!original || /^N\/D$/i.test(original)) {
            return 'N/D';
        }


        const normalized = normalizeText(original)
            .replace(/\s+/g, '');


        const match = normalized.match(
            /^([\d.,]+)(k|m|b|mil|mill|millon|millones)?$/i
        );


        if (!match) return original;


        const suffix = (match[2] || '').toLowerCase();
        const numeric = parseAudienceNumber(
            match[1],
            Boolean(suffix)
        );


        if (!Number.isFinite(numeric)) return original;


        const multiplier =
            suffix === 'k' || suffix === 'mil'
                ? 1_000
                : suffix === 'm' ||
                    suffix === 'mill' ||
                    suffix === 'millon' ||
                    suffix === 'millones'
                    ? 1_000_000
                    : suffix === 'b'
                        ? 1_000_000_000
                        : 1;


        const absolute = numeric * multiplier;


        if (absolute >= 1_000_000) {
            return `${formatCompactCount(absolute / 1_000_000)} millones`;
        }


        if (absolute >= 1_000) {
            return `${formatCompactCount(absolute / 1_000)} mil`;
        }


        return String(Math.round(absolute));
    }


    function parseAudienceNumber(value, hasSuffix) {
        let textValue = String(value || '');
        const dots = (textValue.match(/\./g) || []).length;
        const commas = (textValue.match(/,/g) || []).length;


        if (dots && commas) {
            const decimalSeparator =
                textValue.lastIndexOf('.') > textValue.lastIndexOf(',')
                    ? '.'
                    : ',';


            const thousandsSeparator = decimalSeparator === '.'
                ? ','
                : '.';


            textValue = textValue
                .replace(new RegExp(`\\${thousandsSeparator}`, 'g'), '')
                .replace(decimalSeparator, '.');
        } else if (dots > 1 || commas > 1) {
            textValue = textValue.replace(/[.,]/g, '');
        } else if (dots === 1 || commas === 1) {
            const separator = dots === 1 ? '.' : ',';
            const decimals = textValue.split(separator)[1]?.length || 0;


            if (!hasSuffix && decimals === 3) {
                textValue = textValue.replace(separator, '');
            } else {
                textValue = textValue.replace(separator, '.');
            }
        }


        return Number(textValue);
    }


    function formatCompactCount(value) {
        return Number(value.toFixed(2))
            .toString()
            .replace('.', ',');
    }


    function formatMilitaryDate(day, month, year) {
        const monthAbbreviations = {
            '01': 'ENE',
            '02': 'FEB',
            '03': 'MAR',
            '04': 'ABR',
            '05': 'MAY',
            '06': 'JUN',
            '07': 'JUL',
            '08': 'AGO',
            '09': 'SEP',
            '10': 'OCT',
            '11': 'NOV',
            '12': 'DIC'
        };
        const monthKey = String(month).padStart(2, '0');
        const monthText = monthAbbreviations[monthKey];


        if (!monthText) {
            return 'N/D';
        }


        return `${String(day).padStart(2, '0')}${monthText}${year}`;
    }


    /********************************************************************
     * LOCALIZACIÓN DE IDs INTERNOS
     ********************************************************************/


    function findXInternalId(username) {
        if (!username) return 'N/D';


        const identifierFromScripts = findXIdentifierInScripts(username);
        if (identifierFromScripts) return identifierFromScripts;


        const bodyId = isXAboutPage()
            ? rxFirst(getBodyText(), [
                /(?:Account ID|User ID|ID de la cuenta|ID de usuario)\s*:?\s*(\d{5,})/i
            ])
            : '';


        if (bodyId) return bodyId;


        const stateId = findXIdInReactState(username);


        if (stateId) return stateId;


        const attributedElement = Array.from(document.querySelectorAll(
            '[data-user-id],[data-rest-id],[data-userid]'
        )).find(element => {
            const context = element.closest(
                'header,article,[data-testid="UserName"]'
            );


            return containsExactProfileHandle(
                context?.textContent || '',
                username
            );
        });


        if (attributedElement) {
            const attributedId =
                attributedElement.getAttribute('data-user-id') ||
                attributedElement.getAttribute('data-rest-id') ||
                attributedElement.getAttribute('data-userid') ||
                '';


            if (/^\d{5,}$/.test(attributedId)) {
                return attributedId;
            }
        }


        const allUserLinks = Array.from(document.querySelectorAll('a[href]'))
            .map(link => (link.getAttribute('href') || '').match(
                /\/i\/user\/(\d{5,})/i
            ))
            .filter(Boolean);


        if (allUserLinks.length) {
            return allUserLinks[0][1];
        }


        const userLink = Array.from(document.querySelectorAll('a[href]'))
            .filter(link => {
                const context = link.closest('header,article');


                return containsExactProfileHandle(
                    context?.textContent || '',
                    username
                );
            })
            .map(link => (link.getAttribute('href') || '').match(
                /[?&]user_id=(\d{5,})/i
            ))
            .find(Boolean);


        if (userLink?.[1]) return userLink[1];


        const id = findEmbeddedIdForIdentity(
            username,
            [
                'screen_name',
                'screenName',
                'username'
            ],
            [
                'rest_id',
                'id_str',
                'user_id',
                'id'
            ]
        );


        return id || 'N/D';
    }


    function findXIdentifierInScripts(username) {
        const target = clean(username).replace(/^@/, '');
        if (!target) return '';


        const sources = getEmbeddedTextSources();


        for (const source of sources) {
            for (const variant of embeddedSourceVariants(source)) {
                const identifierPattern = /"identifier"\s*:\s*"(\d{5,})"/gi;
                let match;


                while ((match = identifierPattern.exec(variant)) !== null) {
                    const candidateId = match[1];
                    const windowStart = Math.max(0, match.index - 2000);
                    const windowEnd   = Math.min(variant.length, match.index + match[0].length + 2000);
                    const window      = variant.slice(windowStart, windowEnd);


                    if (
                        containsExactProfileHandle(window, target) ||
                        new RegExp(`"(?:screen_name|username|screenName)"\\s*:\\s*"${escapeRegExp(target)}"`, 'i').test(window) ||
                        new RegExp(`(?:x|twitter)\\.com\\/${escapeRegExp(target)}(?:\\/|[?#]|$)`, 'i').test(window)
                    ) {
                        return candidateId;
                    }
                }


                const markerPattern = new RegExp(
                    `"(?:screen_name|username)"\\s*:\\s*"${escapeRegExp(target)}"`,
                    'gi'
                );


                while ((match = markerPattern.exec(variant)) !== null) {
                    const openings = findOpenJsonObjects(variant, match.index);


                    for (let index = openings.length - 1; index >= 0; index -= 1) {
                        const obj = parseJsonObjectAt(variant, openings[index]);


                        if (!obj) continue;


                        const identifier = clean(
                            obj.identifier ||
                            obj.id_str ||
                            obj.rest_id ||
                            ''
                        );


                        if (/^\d{5,}$/.test(identifier)) {
                            return identifier;
                        }
                    }
                }
            }
        }


        return '';
    }


    function findInstagramPageScopedId(username) {
        if (!sameIdentity(getUserFromUrl(), username)) return '';


        const ownerMeta = document.querySelector(
            'meta[property="instapp:owner_user_id"][content],meta[name="instapp:owner_user_id"][content]'
        );
        const ownerId = clean(ownerMeta?.getAttribute('content') || '');


        if (/^\d{5,}$/.test(ownerId)) {
            return ownerId;
        }


        const loggingIds = new Set();
        const profilePageIds = new Set();


        for (const source of getEmbeddedTextSources()) {
            for (const variant of embeddedSourceVariants(source)) {
                const loggingPatterns = [
                    /"logging_page_id"\s*:\s*"profilePage_(\d{5,})"/gi,
                    /logging_page_id(?:\\"|&quot;)?\s*[:=]\s*(?:\\"|&quot;|"|')?profilePage_(\d{5,})/gi
                ];


                for (const pattern of loggingPatterns) {
                    for (const match of variant.matchAll(pattern)) {
                        loggingIds.add(match[1]);
                    }
                }


                for (const match of variant.matchAll(
                    /\bprofilePage_(\d{5,})\b/gi
                )) {
                    const nearby = variant.slice(
                        Math.max(0, match.index - 5000),
                        Math.min(
                            variant.length,
                            match.index + match[0].length + 5000
                        )
                    );


                    if (
                        instagramSourceHasExactIdentity(
                            nearby,
                            username
                        )
                    ) {
                        profilePageIds.add(match[1]);
                    }
                }
            }
        }


        if (loggingIds.size === 1) {
            return [...loggingIds][0];
        }


        if (loggingIds.size === 0 && profilePageIds.size === 1) {
            return [...profilePageIds][0];
        }


        return '';
    }


    function instagramSourceHasExactIdentity(source, username) {
        const target = clean(username).replace(/^@/, '');


        if (!target) return false;


        const escaped = escapeRegExp(target);
        const patterns = [
            new RegExp(
                `"username"\\s*:\\s*"${escaped}"`,
                'i'
            ),
            new RegExp(
                `\\\\"username\\\\"\\s*:\\s*\\\\"${escaped}\\\\"`,
                'i'
            ),
            new RegExp(
                `instagram\\.com\\/${escaped}(?:\\/|[?#]|$)`,
                'i'
            ),
            new RegExp(
                `(^|[^a-z0-9._])@${escaped}(?=$|[^a-z0-9._])`,
                'i'
            )
        ];


        return patterns.some(pattern => pattern.test(String(source || '')));
    }


    function findXIdInReactState(username) {
        const normalizedUsername = normalizeText(username);
        const usernameElements = Array.from(document.querySelectorAll(
            '[data-testid="UserName"],a[role="link"]'
        )).filter(element => {
            return normalizeText(element.textContent || '')
                .includes(normalizedUsername);
        });


        const roots = [];


        for (const element of usernameElements.slice(0, 8)) {
            let current = element;


            for (let level = 0; current && level < 6; level += 1) {
                for (const key of Object.getOwnPropertyNames(current)) {
                    if (/^__react(?:Props|Fiber|Container)/i.test(key)) {
                        try {
                            roots.push(current[key]);
                        } catch {}
                    }
                }


                current = current.parentElement;
            }
        }


        const visited = new WeakSet();
        let inspected = 0;
        let found = '';


        function visit(value, depth) {
            if (
                found ||
                !value ||
                typeof value !== 'object' ||
                depth > 16 ||
                inspected >= 25000 ||
                visited.has(value)
            ) {
                return;
            }


            visited.add(value);
            inspected += 1;


            const legacy = value.legacy && typeof value.legacy === 'object'
                ? value.legacy
                : null;


            const candidateUsername = clean(
                value.screen_name ||
                value.screenName ||
                value.username ||
                legacy?.screen_name ||
                ''
            ).replace(/^@/, '');


            if (sameIdentity(candidateUsername, username)) {
                const candidateId = clean(
                    value.rest_id ||
                    value.id_str ||
                    value.user_id ||
                    legacy?.id_str ||
                    ''
                );


                if (/^\d{5,}$/.test(candidateId)) {
                    found = candidateId;
                    return;
                }
            }


            if (Array.isArray(value)) {
                for (const item of value) visit(item, depth + 1);
                return;
            }


            for (const key of Object.keys(value)) {
                if (
                    key === 'return' ||
                    key === 'alternate' ||
                    key === '_owner'
                ) {
                    continue;
                }


                try {
                    visit(value[key], depth + 1);
                } catch {}
            }
        }


        for (const root of roots) {
            visit(root, 0);
            if (found) break;
        }


        return found;
    }


    function findInstagramInternalId(username) {
        if (!username) return 'N/D';


        const target = clean(username).replace(/^@/, '');


        const ownerMeta = document.querySelector(
            'meta[property="instapp:owner_user_id"][content],' +
            'meta[name="instapp:owner_user_id"][content]'
        );
        const metaId = clean(ownerMeta?.getAttribute('content') || '');
        if (/^\d{5,}$/.test(metaId)) return metaId;


        const sources = getEmbeddedTextSources();


        for (const source of sources) {
            for (const variant of embeddedSourceVariants(source)) {
                const pkPattern = /"pk"\s*:\s*"?(\d{5,})"?/gi;
                let match;


                while ((match = pkPattern.exec(variant)) !== null) {
                    const candidateId = match[1];
                    const windowStart = Math.max(0, match.index - 3000);
                    const windowEnd   = Math.min(variant.length, match.index + match[0].length + 3000);
                    const window      = variant.slice(windowStart, windowEnd);


                    if (instagramSourceHasExactIdentity(window, target)) {
                        return candidateId;
                    }
                }


                const usernameMarker = new RegExp(
                    `"username"\\s*:\\s*"${escapeRegExp(target)}"`,
                    'gi'
                );


                while ((match = usernameMarker.exec(variant)) !== null) {
                    const openings = findOpenJsonObjects(variant, match.index);


                    for (let index = openings.length - 1; index >= 0; index -= 1) {
                        const obj = parseJsonObjectAt(variant, openings[index]);


                        if (!obj) continue;


                        const candidateId = clean(
                            obj.pk ||
                            obj.id ||
                            obj.user_id ||
                            obj.profile_id ||
                            obj.owner?.id ||
                            obj.user?.pk ||
                            obj.user?.id ||
                            ''
                        );


                        if (/^\d{5,}$/.test(candidateId)) {
                            return candidateId;
                        }
                    }
                }


                const profilePagePattern = /\bprofilePage_(\d{5,})\b/gi;


                while ((match = profilePagePattern.exec(variant)) !== null) {
                    const candidateId = match[1];
                    const windowStart = Math.max(0, match.index - 4000);
                    const windowEnd   = Math.min(variant.length, match.index + match[0].length + 4000);
                    const window      = variant.slice(windowStart, windowEnd);


                    if (instagramSourceHasExactIdentity(window, target)) {
                        return candidateId;
                    }
                }


                const ownerPattern = /"owner"\s*:\s*\{[^}]{0,200}"id"\s*:\s*"(\d{5,})"/gi;


                while ((match = ownerPattern.exec(variant)) !== null) {
                    const candidateId = match[1];
                    const windowStart = Math.max(0, match.index - 3000);
                    const windowEnd   = Math.min(variant.length, match.index + match[0].length + 3000);
                    const window      = variant.slice(windowStart, windowEnd);


                    if (instagramSourceHasExactIdentity(window, target)) {
                        return candidateId;
                    }
                }
            }
        }


        const embeddedId = findEmbeddedIdForIdentity(
            username,
            ['username'],
            ['id', 'pk', 'user_id', 'profile_id']
        );


        if (embeddedId) return embeddedId;


        return findInstagramPageScopedId(username) || 'N/D';
    }


    function findThreadsInternalId(username) {
        if (!username) return 'N/D';


        return findEmbeddedIdForIdentity(
            username,
            [
                'username'
            ],
            [
                'id',
                'pk',
                'user_id'
            ]
        ) || 'N/D';
    }


    function getLoggedInFacebookUserId() {
        const cookieMatch = document.cookie.match(/(?:^|;\s*)c_user=(\d{5,})/);
        if (cookieMatch?.[1]) return cookieMatch[1];


        const html = getHtml();
        const initMatch = html.match(/"CurrentUserInitialData"\s*,\s*\[\]\s*,\s*\{[^}]*"USER_ID"\s*:\s*"(\d{5,})"/i);
        return initMatch?.[1] || '';
    }


    function findFacebookInternalId(type) {
        const mySessionId = getLoggedInFacebookUserId();


        // 1. URL directa (casos inequívocos)
        try {
            const url = new URL(location.href);


            if (type === 'GRUPO') {
                const groupMatch =
                    url.pathname.match(/^\/groups\/(\d{5,})/i);


                if (groupMatch?.[1]) {
                    return groupMatch[1];
                }
            }


            const queryId = url.searchParams.get('id');


            if (
                type !== 'GRUPO' &&
                url.pathname.startsWith('/profile.php') &&
                queryId &&
                /^\d{5,}$/.test(queryId)
            ) {
                return queryId;
            }


            const peopleId = decodeURIComponent(url.pathname).match(
                /^\/people\/[^/]+\/(\d{5,})/i
            )?.[1];


            if (peopleId) {
                return peopleId;
            }
        } catch {}


        // 2. Metadatos de la página (canonical, og:url, al:android:url, al:ios:url)
        const metadataId = findFacebookPageMetadataId(type, mySessionId);
        if (metadataId) return metadataId;


        // 3. Búsqueda en el código fuente (Comet routes, GraphQL, userID, delegate_page_id)
        const sourceId = findFacebookUserIdInSource(type, mySessionId);
        if (sourceId) return sourceId;


        // 4. Enlaces del DOM propios de la cabecera del perfil (álbumes pb.<ID>, fotos, profile.php?id=)
        if (type !== 'GRUPO') {
            const domId = findFacebookIdInDomLinks(mySessionId);
            if (domId) return domId;
        }


        // 5. JSON embebido genérico asociado al nombre/vanity
        const identity = type === 'GRUPO'
            ? getFacebookGroupIdFromUrl() || getFacebookIdentity()
            : getFacebookProfileUsername();


        if (!identity) return 'N/D';


        const embeddedId = findEmbeddedIdForIdentity(
            identity,
            type === 'GRUPO'
                ? ['username', 'vanity', 'slug', 'group_name']
                : ['username', 'vanity', 'vanity_name', 'userVanity'],
            type === 'GRUPO'
                ? ['groupID', 'group_id', 'entity_id', 'id']
                : [
                    'profile_id',
                    'userID',
                    'actorID',
                    'delegate_page_id',
                    'pageID',
                    'entity_id',
                    'id'
                ]
        );


        if (embeddedId && embeddedId !== mySessionId) {
            return embeddedId;
        }


        return 'N/D';
    }


    function findFacebookUserIdInSource(type, mySessionId = '') {
        const sources = getEmbeddedTextSources();


        const profileIdKeys = [
            '"userID"',
            '"UserID"',
            '"user_id"',
            '"profile_id"',
            '"delegate_page_id"',
            '"pageID"',
            '"actorID"',
            '"ownerID"',
            '"owner_id"',
            '"target_id"',
            '"selected_item_id"'
        ];


        const groupIdKeys = [
            '"groupID"',
            '"GroupID"',
            '"group_id"',
            '"entity_id"'
        ];


        const targetKeys = type === 'GRUPO' ? groupIdKeys : profileIdKeys;


        // 1ª pasada: coincidencia estricta con el nombre/vanity o bloque Comet de ruta del perfil
        for (const source of sources) {
            for (const variant of embeddedSourceVariants(source)) {
                // Patrones específicos de ruta Comet de Facebook (muy alta fiabilidad incluso sin vanity cerca)
                if (type !== 'GRUPO') {
                    const cometRoutePatterns = [
                        /"tracePolicy"\s*:\s*"comet\.profile\.[^"]+"[\s\S]{0,1200}?"(?:userID|profile_id|delegate_page_id|entity_id)"\s*:\s*"?(\d{5,})"?/gi,
                        /"(?:userID|profile_id|delegate_page_id|entity_id)"\s*:\s*"?(\d{5,})"?[\s\S]{0,1200}?"tracePolicy"\s*:\s*"comet\.profile\.[^"]+"/gi,
                        /"profile_owner"\s*:\s*\{[^}]{0,250}?"id"\s*:\s*"(\d{5,})"/gi,
                        /"ProfileCometHeader[^"]*"[\s\S]{0,1000}?"(?:userID|profileID|id)"\s*:\s*"(\d{5,})"/gi
                    ];


                    for (const rtPattern of cometRoutePatterns) {
                        let rtMatch;
                        while ((rtMatch = rtPattern.exec(variant)) !== null) {
                            const candidate = rtMatch[1];
                            if (candidate && candidate !== mySessionId) {
                                return candidate;
                            }
                        }
                    }
                }


                for (const key of targetKeys) {
                    const pattern = new RegExp(
                        `${escapeRegExp(key)}\\s*:\\s*"?(\\d{5,})"?`,
                        'g'
                    );


                    let match;


                    while ((match = pattern.exec(variant)) !== null) {
                        const candidateId = match[1];
                        if (candidateId === mySessionId) continue;


                        const windowStart = Math.max(0, match.index - 3500);
                        const windowEnd   = Math.min(
                            variant.length,
                            match.index + match[0].length + 3500
                        );
                        const window = variant.slice(windowStart, windowEnd);


                        if (facebookSourceMatchesCurrentProfile(window, type)) {
                            return candidateId;
                        }
                    }
                }
            }
        }


        return '';
    }


    function findFacebookIdInDomLinks(mySessionId = '') {
        const scope = document.querySelector('[role="main"]') || document.body;
        if (!scope) return '';


        const links = Array.from(scope.querySelectorAll('a[href]'));


        for (const link of links) {
            if (link.closest('[role="article"]')) continue;
            const href = link.getAttribute('href') || '';


            // Enlaces de álbumes de perfil: set=pb.<USER_ID>.-2207520000 o set=a.<ID>.<ID>.<USER_ID>
            const pbMatch = href.match(/[?&]set=pb\.(\d{5,})\./i);
            if (pbMatch?.[1] && pbMatch[1] !== mySessionId) {
                return pbMatch[1];
            }


            const profilePhpMatch = href.match(/profile\.php\?id=(\d{5,})/i);
            if (profilePhpMatch?.[1] && profilePhpMatch[1] !== mySessionId) {
                return profilePhpMatch[1];
            }
        }


        return '';
    }


    function facebookSourceMatchesCurrentProfile(source, type) {
        try {
            const url = new URL(location.href);


            if (type === 'GRUPO') {
                const groupId = getFacebookGroupIdFromUrl();


                if (groupId) {
                    if (/^\d{5,}$/.test(groupId)) {
                        return source.includes(groupId);
                    }


                    return new RegExp(
                        `"(?:vanity|slug|username|group_name)"\\s*:\\s*"${escapeRegExp(groupId)}"`,
                        'i'
                    ).test(source);
                }
            }


            const queryId = url.searchParams.get('id');


            if (url.pathname.startsWith('/profile.php') && queryId) {
                return source.includes(queryId);
            }


            const username = getFacebookProfileUsername();


            if (username) {
                const escaped = escapeRegExp(username);
                if (
                    new RegExp(`"(?:vanity|vanity_name|userVanity|username|login|profile_url|url)"\\s*:\\s*"[^"]*${escaped}[^"]*"`, 'i').test(source) ||
                    source.includes(`facebook.com/${username}`) ||
                    source.includes(`/${username}"`) ||
                    new RegExp(`(?:^|[^a-z0-9._])${escaped}(?:$|[^a-z0-9._])`, 'i').test(source)
                ) {
                    return true;
                }
            }


            // También aceptar si el bloque contiene el nombre visible exacto del perfil (H1)
            const displayName = extractFacebookDisplayName();
            if (displayName && displayName.length >= 3) {
                const escapedName = escapeRegExp(displayName);
                if (new RegExp(`"name"\\s*:\\s*"${escapedName}"`, 'i').test(source)) {
                    return true;
                }
            }


            return false;
        } catch {
            return false;
        }
    }


    function findFacebookPageMetadataId(type, mySessionId = '') {
        const values = [
            location.href,
            ...Array.from(document.querySelectorAll(
                'link[rel="canonical"][href],meta[property="og:url"][content],meta[property="al:ios:url"][content],meta[property="al:android:url"][content]'
            )).map(element => {
                return element.getAttribute('href') ||
                    element.getAttribute('content') ||
                    '';
            })
        ];


        const patterns = type === 'GRUPO'
            ? [
                /\/groups\/(\d{5,})/i,
                /(?:fb:\/\/group\/|group_id[=/:]|groupID[=/:])(\d{5,})/i
            ]
            : [
                /profile\.php\?[^#]*\bid=(\d{5,})/i,
                /\/people\/[^/]+\/(\d{5,})/i,
                /(?:fb:\/\/(?:profile|page)\/(?:\?id=)?|profile_id[=/:]|page_id[=/:])(\d{5,})/i
            ];


        for (const value of values) {
            const id = rxFirst(String(value || ''), patterns);


            if (id && id !== mySessionId) return id;
        }


        return '';
    }


    function containsExactProfileHandle(source, username) {
        const handle = clean(username).replace(/^@/, '');


        if (!handle) return false;


        return new RegExp(
            `(^|[^a-z0-9_])@${escapeRegExp(handle)}(?=$|[^a-z0-9_])`,
            'i'
        ).test(String(source || ''));
    }


    function findEmbeddedIdForIdentity(
        username,
        identityKeys,
        idKeys
    ) {
        const target = clean(username).replace(/^@/, '');


        if (!target) return '';


        for (const root of getEmbeddedJsonRoots()) {
            const id = findEntityIdInObjectTree(
                root,
                target,
                identityKeys,
                idKeys
            );


            if (id) return id;
        }


        const encodedTarget = JSON.stringify(target).slice(1, -1);


        for (const source of getEmbeddedTextSources()) {
            for (const variant of embeddedSourceVariants(source)) {
                for (const identityKey of identityKeys) {
                    const marker = new RegExp(
                        `"${escapeRegExp(identityKey)}"\\s*:\\s*"${escapeRegExp(encodedTarget)}"`,
                        'gi'
                    );
                    let match;
                    let inspected = 0;


                    while ((match = marker.exec(variant)) && inspected < 20) {
                        const openings = findOpenJsonObjects(
                            variant,
                            match.index
                        );


                        for (
                            let index = openings.length - 1;
                            index >= 0;
                            index -= 1
                        ) {
                            const object = parseJsonObjectAt(
                                variant,
                                openings[index]
                            );


                            if (!object) continue;


                            const id = findEntityIdInObjectTree(
                                object,
                                target,
                                identityKeys,
                                idKeys
                            );


                            if (id) return id;
                        }


                        inspected += 1;
                    }
                }
            }
        }


        return '';
    }


    function getEmbeddedJsonRoots() {
        const roots = [];


        for (const source of getEmbeddedTextSources()) {
            for (const variant of embeddedSourceVariants(source)) {
                const trimmed = variant.trim();


                if (!/^[\[{]/.test(trimmed)) continue;


                try {
                    const parsed = JSON.parse(trimmed);


                    if (parsed && typeof parsed === 'object') {
                        roots.push(parsed);
                    }
                } catch {}
            }
        }


        return roots;
    }


    function embeddedSourceVariants(source) {
        const original = String(source || '');
        const variants = [original];


        if (original.includes('&quot;')) {
            variants.push(decodeHtml(original));
        }


        if (original.includes('\\"')) {
            variants.push(original.replace(/\\"/g, '"'));
        }


        if (original.includes('%22')) {
            try {
                variants.push(decodeURIComponent(original));
            } catch {}
        }


        return [...new Set(variants)];
    }


    function findEntityIdInObjectTree(
        root,
        username,
        identityKeys,
        idKeys
    ) {
        const visited = new WeakSet();
        let inspected = 0;


        function visit(value, depth) {
            if (
                !value ||
                typeof value !== 'object' ||
                depth > 18 ||
                inspected >= 40000 ||
                visited.has(value)
            ) {
                return '';
            }


            visited.add(value);
            inspected += 1;


            if (objectHasExactIdentity(value, username, identityKeys)) {
                const directId = pickNumericObjectValue(value, idKeys);


                if (directId) return directId;
            }


            for (const key of [
                'legacy',
                'user',
                'profile',
                'owner',
                'account',
                'node'
            ]) {
                const nested = value[key];


                if (
                    nested &&
                    typeof nested === 'object' &&
                    objectHasExactIdentity(
                        nested,
                        username,
                        identityKeys
                    )
                ) {
                    const nestedId =
                        pickNumericObjectValue(nested, idKeys) ||
                        pickNumericObjectValue(value, idKeys);


                    if (nestedId) return nestedId;
                }
            }


            const children = Array.isArray(value)
                ? value
                : Object.values(value);


            for (const child of children) {
                const found = visit(child, depth + 1);


                if (found) return found;
            }


            return '';
        }


        return visit(root, 0);
    }


    function objectHasExactIdentity(object, username, identityKeys) {
        if (!object || typeof object !== 'object') return false;


        const allowedKeys = new Set(
            identityKeys.map(key => key.toLowerCase())
        );


        return Object.entries(object).some(([key, value]) => {
            if (!allowedKeys.has(key.toLowerCase())) return false;
            if (typeof value !== 'string') return false;


            return sameIdentity(
                normalizeEmbeddedIdentity(value),
                username
            );
        });
    }


    function normalizeEmbeddedIdentity(value) {
        const cleaned = clean(value).replace(/^@/, '');


        try {
            const url = new URL(cleaned, location.origin);
            const pathIdentity = decodeURIComponent(url.pathname)
                .split('/')
                .filter(Boolean)[0]
                ?.replace(/^@/, '');


            return pathIdentity || cleaned;
        } catch {
            return cleaned;
        }
    }


    function pickNumericObjectValue(object, keys) {
        if (!object || typeof object !== 'object') return '';


        const allowedKeys = new Set(keys.map(key => key.toLowerCase()));


        for (const [key, value] of Object.entries(object)) {
            if (!allowedKeys.has(key.toLowerCase())) continue;


            if (
                typeof value === 'number' &&
                !Number.isSafeInteger(value)
            ) {
                continue;
            }


            const candidate = clean(value);


            if (/^\d{5,}$/.test(candidate)) {
                return candidate;
            }
        }


        return '';
    }


    function getEmbeddedTextSources() {
        const scripts = Array.from(
            document.querySelectorAll('script')
        );


        const sources = scripts
            .map(script => script.textContent || '')
            .filter(textValue => textValue.length > 20);


        sources.push(getHtml());


        return sources;
    }


    /********************************************************************
     * URLS Y TIPOS DE PÁGINA
     ********************************************************************/


    function detectPlatform() {
        const hostname = location.hostname
            .replace(/^www\./, '')
            .toLowerCase();


        if (hostname === 'tiktok.com') return 'TikTok';


        if (
            hostname === 'x.com' ||
            hostname === 'twitter.com'
        ) {
            return 'Twitter/X';
        }


        if (
            hostname === 'facebook.com' ||
            hostname === 'm.facebook.com'
        ) {
            return 'Facebook';
        }


        if (hostname === 'instagram.com') {
            return 'Instagram';
        }


        if (
            hostname === 'threads.com' ||
            hostname === 'threads.net'
        ) {
            return 'Threads';
        }


        return '';
    }


    function getUserFromUrl() {
        try {
            const url = new URL(location.href);
            const hostname = url.hostname
                .replace(/^www\./, '')
                .toLowerCase();


            const segments = decodeURIComponent(url.pathname)
                .split('/')
                .filter(Boolean);


            if (hostname === 'tiktok.com') {
                const match = url.pathname.match(
                    /^\/@([^/?#]+)/
                );


                return match?.[1] || '';
            }


            if (
                hostname === 'x.com' ||
                hostname === 'twitter.com'
            ) {
                const first = segments[0] || '';


                const reserved = [
                    'home',
                    'explore',
                    'notifications',
                    'messages',
                    'search',
                    'settings',
                    'compose',
                    'i'
                ];


                return reserved.includes(first.toLowerCase())
                    ? ''
                    : first.replace(/^@/, '');
            }


            if (hostname === 'instagram.com') {
                const first = segments[0] || '';


                const reserved = [
                    'p',
                    'reel',
                    'stories',
                    'explore',
                    'accounts',
                    'direct'
                ];


                return reserved.includes(first.toLowerCase())
                    ? ''
                    : first.replace(/^@/, '');
            }


            if (
                hostname === 'threads.com' ||
                hostname === 'threads.net'
            ) {
                const match = url.pathname.match(
                    /^\/@([^/?#]+)/
                );


                return match?.[1] || '';
            }


            if (
                hostname === 'facebook.com' ||
                hostname === 'm.facebook.com'
            ) {
                return getFacebookIdentity();
            }


            return '';
        } catch {
            return '';
        }
    }


    function getFacebookObjectType() {
        return /^\/groups\//i.test(location.pathname)
            ? 'GRUPO'
            : 'PERFIL';
    }


    function getFacebookGroupIdFromUrl() {
        try {
            return new URL(location.href).pathname
                .match(/^\/groups\/([^/?#]+)/i)?.[1] || '';
        } catch {
            return '';
        }
    }


    function getFacebookIdentity() {
        try {
            const url = new URL(location.href);
            const segments = decodeURIComponent(url.pathname)
                .split('/')
                .filter(Boolean);


            const groupMatch =
                url.pathname.match(/^\/groups\/([^/?#]+)/i);


            if (groupMatch?.[1]) {
                return groupMatch[1];
            }


            const queryId = url.searchParams.get('id');


            if (
                url.pathname.startsWith('/profile.php') &&
                queryId
            ) {
                return queryId;
            }


            if (segments[0]?.toLowerCase() === 'people') {
                const numericId = [...segments]
                    .reverse()
                    .find(segment => /^\d{5,}$/.test(segment));


                return numericId || segments[1] || '';
            }


            return segments[0] || '';
        } catch {
            return '';
        }
    }


    function getFacebookProfileUrlName() {
        if (getFacebookObjectType() === 'GRUPO') return '';


        try {
            const url = new URL(location.href);
            const queryId = url.searchParams.get('id');


            if (url.pathname.startsWith('/profile.php') && queryId) {
                return queryId;
            }


            const segments = decodeURIComponent(url.pathname)
                .split('/')
                .filter(Boolean);


            if (segments[0]?.toLowerCase() === 'people') {
                return segments[1] || getFacebookIdentity();
            }


            return getFacebookProfileUsername();
        } catch {
            return '';
        }
    }


    function getFacebookProfileUsername() {
        if (getFacebookObjectType() === 'GRUPO') {
            return '';
        }


        const identity = getFacebookIdentity();


        const reserved = [
            'profile.php',
            'people',
            'home.php',
            'watch',
            'marketplace',
            'groups',
            'events',
            'friends',
            'messages',
            'login'
        ];


        if (
            !identity ||
            reserved.includes(identity.toLowerCase())
        ) {
            return '';
        }


        return identity;
    }


    function isXAboutPage() {
        return /\/about\/?$/i.test(location.pathname);
    }


    function isFacebookAboutPage() {
        try {
            const url = new URL(location.href);


            return (
                /\/about(?:_profile_transparency)?\/?$/i.test(url.pathname) ||
                /^about/i.test(url.searchParams.get('sk') || '')
            );
        } catch {
            return false;
        }
    }


    function facebookAboutUrl() {
        try {
            const url = new URL(location.href);


            if (getFacebookObjectType() === 'GRUPO') {
                const match =
                    url.pathname.match(/^\/groups\/([^/?#]+)/i);


                if (!match?.[1]) return '';


                return `${url.origin}/groups/${match[1]}/about`;
            }


            if (url.pathname.startsWith('/profile.php')) {
                const id = url.searchParams.get('id');


                if (!id) return '';


                return `${url.origin}/profile.php?id=${encodeURIComponent(id)}&sk=about_profile_transparency`;
            }


            const peopleMatch = decodeURIComponent(url.pathname).match(
                /^\/people\/[^/]+\/(\d{5,})/i
            );


            if (peopleMatch?.[1]) {
                return `${url.origin}/profile.php?id=${encodeURIComponent(peopleMatch[1])}&sk=about_profile_transparency`;
            }


            const username = getFacebookProfileUsername();


            if (!username) return '';


            return `${url.origin}/${encodeURIComponent(username)}/about_profile_transparency`;
        } catch {
            return '';
        }
    }


    function canonicalProfileUrl() {
        try {
            const url = new URL(location.href);


            url.hash = '';


            if (PLATFORM === 'Twitter/X') {
                const username = getUserFromUrl();


                return username
                    ? `https://x.com/${encodeURIComponent(username)}`
                    : url.href;
            }


            if (PLATFORM === 'Instagram') {
                const username = getUserFromUrl();


                return username
                    ? `https://www.instagram.com/${encodeURIComponent(username)}/`
                    : url.href;
            }


            if (PLATFORM === 'TikTok') {
                const username = getUserFromUrl();


                return username
                    ? `https://www.tiktok.com/@${encodeURIComponent(username)}`
                    : url.href;
            }


            if (PLATFORM === 'Threads') {
                const username = getUserFromUrl();


                return username
                    ? `https://www.threads.com/@${encodeURIComponent(username)}`
                    : url.href;
            }


            if (PLATFORM === 'Facebook') {
                if (getFacebookObjectType() === 'GRUPO') {
                    const identity = getFacebookIdentity();


                    return `${url.origin}/groups/${identity}`;
                }


                const id = url.searchParams.get('id');


                if (url.pathname.startsWith('/profile.php') && id) {
                    return `${url.origin}/profile.php?id=${encodeURIComponent(id)}`;
                }


                const username = getFacebookProfileUsername();


                return username
                    ? `${url.origin}/${encodeURIComponent(username)}`
                    : url.href;
            }


            return url.href;
        } catch {
            return location.href;
        }
    }


    function toProfileUrl(input) {
        let value = clean(input);


        if (!value) return '';


        if (/^https?:\/\//i.test(value)) {
            return value;
        }


        value = value.replace(/^@/, '');


        switch (PLATFORM) {
            case 'TikTok':
                return `https://www.tiktok.com/@${encodeURIComponent(value)}`;


            case 'Twitter/X':
                return `https://x.com/${encodeURIComponent(value)}`;


            case 'Instagram':
                return `https://www.instagram.com/${encodeURIComponent(value)}/`;


            case 'Threads':
                return `https://www.threads.com/@${encodeURIComponent(value)}`;


            case 'Facebook':
                if (/^\d{5,}$/.test(value)) {
                    return `https://www.facebook.com/profile.php?id=${encodeURIComponent(value)}`;
                }


                return `https://www.facebook.com/${encodeURIComponent(value)}`;


            default:
                return '';
        }
    }


    function sameProfileTarget(targetUrl, currentUrl) {
        const target = extractTargetIdentity(targetUrl);
        const current = extractTargetIdentity(currentUrl);


        if (!target || !current) return false;


        return sameIdentity(target, current);
    }


    function extractTargetIdentity(rawUrl) {
        try {
            const url = new URL(rawUrl, location.origin);
            const path = decodeURIComponent(url.pathname);


            if (PLATFORM === 'TikTok') {
                return path.match(/^\/@([^/?#]+)/)?.[1] || '';
            }


            if (PLATFORM === 'Twitter/X') {
                return path.split('/').filter(Boolean)[0] || '';
            }


            if (PLATFORM === 'Instagram') {
                return path.split('/').filter(Boolean)[0] || '';
            }


            if (PLATFORM === 'Threads') {
                return path.match(/^\/@([^/?#]+)/)?.[1] || '';
            }


            if (PLATFORM === 'Facebook') {
                const group =
                    path.match(/^\/groups\/([^/?#]+)/i);


                if (group?.[1]) {
                    return `group:${group[1]}`;
                }


                if (path.startsWith('/profile.php')) {
                    return `profile:${url.searchParams.get('id') || ''}`;
                }


                return `profile:${path.split('/').filter(Boolean)[0] || ''}`;
            }


            return '';
        } catch {
            return '';
        }
    }


    function sameIdentity(first, second) {
        return normalizeText(first) === normalizeText(second);
    }


    /********************************************************************
     * DETECCIÓN DE ERRORES DE PLATAFORMA
     ********************************************************************/


    function detectBlock() {
        const content = normalizeText(
            `${document.title}\n${getBodyText()}`
        );


        return /temporarily blocked|too many requests|rate limit|try again later|unusual traffic|captcha|please wait a few minutes|demasiadas solicitudes|intentalo de nuevo mas tarde|actividad inusual|verifica que eres tu|verificacion de seguridad|hemos limitado la frecuencia/.test(content);
    }


    function detectNotFound() {
        const content = normalizeText(
            `${document.title}\n${getBodyText()}`
        );


        switch (PLATFORM) {
            case 'TikTok': {
                if (document.querySelector('[data-e2e="error-page"]')) {
                    return true;
                }


                const bundle = findTikTokUserBundle(getUserFromUrl() || '');
                if (bundle.notFound) {
                    return true;
                }


                return /couldn.t find this account|account doesn.t exist|account not found|usuario no encontrado|cuenta no encontrada|no se ha podido encontrar esta cuenta|no se pudo encontrar esta cuenta|esta cuenta no existe|cuenta suspendida|this account was banned|account banned/.test(content);
            }


            case 'Twitter/X':
                return /this account doesn.t exist|try searching for another|esta cuenta no existe|account suspended|cuenta suspendida/.test(content);


            case 'Instagram':
                return /sorry, this page isn.t available|esta pagina no esta disponible|user not found|usuario no encontrado/.test(content);


            case 'Facebook':
                return /this content isn.t available|this page isn.t available|esta pagina no esta disponible|este contenido no esta disponible|perfil no disponible/.test(content);


            case 'Threads':
                return /this profile isn.t available|profile not found|perfil no disponible|esta pagina no esta disponible/.test(content);


            default:
                return false;
        }
    }


    function buildUnavailableRow(username, exists, observations) {
        const missingValue = normalizeText(exists) === 'no'
            ? '-'
            : 'N/D';


        return {
            USUARIO: username || missingValue,
            'NOMBRE VISIBLE': missingValue,
            'ID interno': missingValue,
            'Nº SEGUIDORES': missingValue,
            'Nº SEGUIDOS': missingValue,
            'TIPO DE OBJETO':
                PLATFORM === 'Facebook'
                    ? getFacebookObjectType()
                    : 'PERFIL',
            MÉTRICA:
                PLATFORM === 'Facebook' &&
                getFacebookObjectType() === 'GRUPO'
                    ? 'MIEMBROS'
                    : PLATFORM === 'Facebook'
                        ? 'AMIGOS'
                        : 'SEGUIDORES',
            PLATAFORMA: PLATFORM,
            'FECHA DE CREACIÓN DE CUENTA': missingValue,
            'UBICACIÓN DE LA CUENTA': missingValue,
            'ÚLTIMO CAMBIO DE USUARIO': missingValue,
            'ÚLTIMA ACTUALIZACIÓN DEL PERFIL': missingValue,
            'Nº CAMBIOS DE USUARIO': missingValue,
            EXISTE: exists,
            URL: canonicalProfileUrl(),
            OBSERVACIONES: observations,
            'GRUPO PÚBLICO': missingValue,
            'GRUPO VISIBLE': missingValue
        };
    }


    /********************************************************************
     * HELPERS DE TEXTO Y DOM
     ********************************************************************/


    function $(selector) {
        return document.querySelector(selector);
    }


    function text(selector) {
        const element = document.querySelector(selector);


        return clean(
            element?.innerText ||
            element?.textContent ||
            ''
        );
    }


    function metaContent(selector) {
        const element = document.querySelector(selector);


        return clean(
            element?.getAttribute('content') ||
            ''
        );
    }


    function getBodyText() {
        return document.body?.innerText || '';
    }


    function getHtml() {
        return document.documentElement?.innerHTML || '';
    }


    function clean(value) {
        return String(value ?? '')
            .replace(/\u00a0/g, ' ')
            .replace(/\s+/g, ' ')
            .trim();
    }


    function normalizeText(value) {
        return clean(value)
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .toLowerCase();
    }


    function rxFirst(source, patterns) {
        const textValue = String(source || '');


        for (const pattern of patterns) {
            const match = textValue.match(pattern);


            if (match?.[1]) {
                return decodeHtml(match[1]);
            }
        }


        return '';
    }


    function extractMetric(source, patterns) {
        return clean(rxFirst(source, patterns)) || 'N/D';
    }


    function extractFollowingMetric(source) {
        const value = rxFirst(source, [
            /(?:^|[\s•·|,\n\r])([\d][\d.,\s]*(?:K|M|B|mil|millones)?)\s+(?:Following|Siguiendo|Seguidos|Seguidas)\b/i,
            /(?:Following|Siguiendo|Seguidos|Seguidas)\s*:?\s*([\d][\d.,\s]*(?:K|M|B|mil|millones)?)/i
        ]);


        return clean(value);
    }


    function extractMetricFromLines(source, patterns, exclusions = []) {
        const lines = String(source || '')
            .split(/\r?\n/)
            .map(clean)
            .filter(Boolean);


        for (const line of lines) {
            if (
                exclusions.some(pattern => pattern.test(line))
            ) {
                continue;
            }


            for (const pattern of patterns) {
                const match = line.match(pattern);


                if (match?.[1]) {
                    return clean(match[1]);
                }
            }
        }


        return 'N/D';
    }


    function extractLabeledValue(source, labels) {
        const lines = String(source || '')
            .split(/\r?\n/)
            .map(clean)
            .filter(Boolean);


        for (let index = 0; index < lines.length; index += 1) {
            const line = lines[index];


            for (const label of labels) {
                const match = line.match(label);


                if (!match) continue;


                if (match[1]) {
                    return clean(match[1]);
                }


                const trailing = clean(
                    line
                        .slice(match.index + match[0].length)
                        .replace(/^[:\-–—\s]+/, '')
                );


                if (
                    trailing &&
                    trailing.length <= 120 &&
                    !isUiNoise(trailing)
                ) {
                    return trailing;
                }


                for (
                    let nextIndex = index + 1;
                    nextIndex <= index + 3 &&
                    nextIndex < lines.length;
                    nextIndex += 1
                ) {
                    const candidate = lines[nextIndex];


                    if (
                        candidate &&
                        candidate.length <= 120 &&
                        !isUiNoise(candidate)
                    ) {
                        return candidate;
                    }
                }
            }
        }


        return '';
    }


    function extractCountAroundLabel(source, labels) {
        const lines = String(source || '')
            .split(/\r?\n/)
            .map(clean)
            .filter(Boolean);


        const directText = clean(source);


        for (const label of labels) {
            const normalizedLabel = label.source;


            const afterPattern = new RegExp(
                `${normalizedLabel}\\s*:?\\s*(\\d{1,5})`,
                'i'
            );


            const afterMatch = directText.match(afterPattern);


            if (afterMatch?.[1]) {
                return afterMatch[1];
            }


            const beforePattern = new RegExp(
                `(\\d{1,5})\\s*${normalizedLabel}`,
                'i'
            );


            const beforeMatch = directText.match(beforePattern);


            if (beforeMatch?.[1]) {
                return beforeMatch[1];
            }
        }


        for (let index = 0; index < lines.length; index += 1) {
            if (
                !labels.some(label => label.test(lines[index]))
            ) {
                continue;
            }


            const nearby = lines
                .slice(index, index + 5)
                .join(' ');


            if (
                /no former usernames|ningun nombre de usuario anterior|ningún nombre de usuario anterior/i.test(nearby)
            ) {
                return '0';
            }


            const match = nearby.match(/\b(\d{1,5})\b/);


            if (match?.[1]) {
                return match[1];
            }
        }


        return '';
    }


    function isUiNoise(value) {
        return /^(close|cerrar|back|volver|more|más|mas|learn more|más información|mas informacion|about this account|acerca de esta cuenta)$/i.test(
            clean(value)
        );
    }


    function cleanAccountLocation(value) {
        const cleaned = clean(value);


        if (!cleaned) return 'N/D';


        return cleaned
            .replace(
                /^(Account based in|Cuenta basada en|País donde se encuentra la cuenta|Ubicación de la cuenta)\s*:?\s*/i,
                ''
            )
            .trim() || 'N/D';
    }


    function findClickableByText(
        labels,
        root = document,
        exactOnly = false
    ) {
        const normalizedLabels = labels.map(normalizeText);


        const elements = Array.from(
            root.querySelectorAll(
                'button,[role="button"],[role="menuitem"],a,div[tabindex="0"]'
            )
        );


        let partialMatch = null;


        for (const element of elements) {
            if (!isVisible(element)) continue;


            const value = clean(
                element.innerText ||
                element.textContent ||
                ''
            );


            if (!value || value.length > 140) continue;


            const normalizedValue = normalizeText(value);


            if (normalizedLabels.includes(normalizedValue)) {
                return element;
            }


            if (
                !exactOnly &&
                !partialMatch &&
                normalizedLabels.some(label => {
                    return normalizedValue.includes(label);
                })
            ) {
                partialMatch = element;
            }
        }


        return partialMatch;
    }


    function safeClick(element) {
        if (!element) return false;


        try {
            element.scrollIntoView({
                block: 'center',
                inline: 'center'
            });
        } catch {}


        try {
            element.focus();
        } catch {}


        try {
            element.click();
            return true;
        } catch {}


        try {
            element.dispatchEvent(
                new MouseEvent('mousedown', {
                    bubbles: true,
                    cancelable: true,
                    view: window
                })
            );


            element.dispatchEvent(
                new MouseEvent('mouseup', {
                    bubbles: true,
                    cancelable: true,
                    view: window
                })
            );


            element.dispatchEvent(
                new MouseEvent('click', {
                    bubbles: true,
                    cancelable: true,
                    view: window
                })
            );


            return true;
        } catch {}


        try {
            element.dispatchEvent(
                new KeyboardEvent('keydown', {
                    key: 'Enter',
                    code: 'Enter',
                    bubbles: true,
                    cancelable: true
                })
            );


            element.dispatchEvent(
                new KeyboardEvent('keyup', {
                    key: 'Enter',
                    code: 'Enter',
                    bubbles: true,
                    cancelable: true
                })
            );


            return true;
        } catch {}


        return false;
    }


    function isVisible(element) {
        if (!element) return false;


        const style = getComputedStyle(element);
        const rect = element.getBoundingClientRect();


        return (
            style.display !== 'none' &&
            style.visibility !== 'hidden' &&
            Number(style.opacity || 1) !== 0 &&
            rect.width > 0 &&
            rect.height > 0
        );
    }


    async function waitFor(callback, timeout = 5000, interval = 200) {
        const start = Date.now();


        while (Date.now() - start < timeout) {
            try {
                const result = callback();


                if (result) {
                    return result;
                }
            } catch {}


            await sleep(interval);
        }


        return null;
    }


    function waitForDom(callback, timeout = 8000) {
        return new Promise(resolve => {
            let finished = false;
            let observer = null;


            const finish = value => {
                if (finished) return;


                finished = true;
                observer?.disconnect();
                clearTimeout(timer);
                resolve(value || null);
            };


            const check = () => {
                try {
                    const result = callback();


                    if (result) {
                        finish(result);
                    }
                } catch {}
            };


            const timer = setTimeout(() => {
                check();
                finish(null);
            }, timeout);


            observer = new MutationObserver(check);
            observer.observe(document.documentElement, {
                childList: true,
                subtree: true,
                attributes: true,
                attributeFilter: [
                    'aria-expanded',
                    'aria-modal',
                    'aria-label'
                ]
            });


            check();
        });
    }


    async function waitForPage() {
        const start = Date.now();


        while (Date.now() - start < 15000) {
            if (document.body && getBodyText().length > 60) {
                break;
            }
            await sleep(350);
        }


        // Espera adicional específica de plataforma para asegurar hidratación del perfil
        await waitFor(() => {
            if (detectNotFound() || detectBlock()) return true;


            switch (PLATFORM) {
                case 'TikTok':
                    return Boolean(
                        document.querySelector('[data-e2e="user-title"]') ||
                        document.querySelector('[data-e2e="followers-count"]') ||
                        document.querySelector('[data-e2e="error-page"]')
                    );
                case 'Facebook':
                    return Boolean(
                        Array.from(document.querySelectorAll('h1')).find(isVisible)
                    );
                case 'Twitter/X':
                    return Boolean(
                        document.querySelector('[data-testid="UserName"]') ||
                        isXAboutPage()
                    );
                case 'Instagram':
                    return Boolean(
                        document.querySelector('main header') ||
                        document.querySelector('meta[property="og:description"]')
                    );
                default:
                    return true;
            }
        }, 5000, 250);
    }


    function parseJsonScript(id) {
        try {
            const script = document.getElementById(id);


            if (!script?.textContent) {
                return null;
            }


            return JSON.parse(script.textContent);
        } catch {
            return null;
        }
    }


    function walkObject(root, visitor) {
        const visited = new WeakSet();
        let visitedCount = 0;
        const maxObjects = 30000;


        function walk(value, depth) {
            if (
                !value ||
                typeof value !== 'object' ||
                depth > 14 ||
                visitedCount >= maxObjects
            ) {
                return;
            }


            if (visited.has(value)) return;


            visited.add(value);
            visitedCount += 1;
            visitor(value);


            if (Array.isArray(value)) {
                for (const item of value) {
                    walk(item, depth + 1);
                }
            } else {
                for (const key of Object.keys(value)) {
                    walk(value[key], depth + 1);
                }
            }
        }


        walk(root, 0);
    }


    function pick(object, keys) {
        if (!object || typeof object !== 'object') {
            return '';
        }


        for (const key of keys) {
            if (
                object[key] !== undefined &&
                object[key] !== null &&
                object[key] !== ''
            ) {
                return String(object[key]);
            }
        }


        return '';
    }


    function mergeRows(...objects) {
        const output = {};


        for (const object of objects) {
            if (!object || typeof object !== 'object') {
                continue;
            }


            for (const [key, value] of Object.entries(object)) {
                if (key === 'OBSERVACIONES') {
                    output[key] = joinNotes(
                        output[key],
                        value
                    );


                    continue;
                }


                if (
                    isUsefulValue(value) ||
                    !(key in output)
                ) {
                    output[key] = value;
                }
            }
        }


        return output;
    }


    function isUsefulValue(value) {
        const cleaned = clean(value);


        return (
            cleaned !== '' &&
            cleaned !== '-' &&
            !/^N\/D$/i.test(cleaned) &&
            !/^undefined$/i.test(cleaned) &&
            !/^null$/i.test(cleaned)
        );
    }


    function joinNotes(...notes) {
        const unique = [];


        for (const note of notes) {
            const cleaned = clean(note);


            if (
                !cleaned ||
                /^N\/D$/i.test(cleaned) ||
                unique.includes(cleaned)
            ) {
                continue;
            }


            unique.push(cleaned);
        }


        return unique.join(' ');
    }


    function rowKey(row) {
        const id = clean(row['ID interno']);


        if (isUsefulValue(id)) {
            return [
                normalizeText(row.PLATAFORMA || PLATFORM),
                normalizeText(row['TIPO DE OBJETO'] || ''),
                `id:${id}`
            ].join('::');
        }


        return [
            normalizeText(row.PLATAFORMA || PLATFORM),
            normalizeText(row['TIPO DE OBJETO'] || ''),
            normalizeText(row.USUARIO || '')
        ].join('::');
    }


    function setStatus(message) {
        const element = $('#falconai-status');


        if (element) {
            element.textContent = message;
        }


        console.log('[Roomba Scraper]', message);
    }


    function updatePanelCounts() {
        const rows = getRows();
        const queue = getQueue();


        const element = $('#falconai-small');


        if (!element) return;


        element.textContent =
            `Filas: ${rows.length} · ` +
            `Cola: ${queue.items.length} · ` +
            `Índice: ${queue.index} · ` +
            `Estado: ${queue.active ? 'activa' : 'pausada'}. ` +
            `Los datos no visibles se registran como N/D; no se infieren ni se sortean verificaciones.`;
    }


    function decodeHtml(value) {
        const textarea = document.createElement('textarea');
        textarea.innerHTML = String(value || '');
        return textarea.value;
    }


    function escapeHtml(value) {
        return String(value ?? '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }


    function escapeRegExp(value) {
        return String(value)
            .replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }


    function sleep(milliseconds) {
        return new Promise(resolve => {
            setTimeout(resolve, milliseconds);
        });
    }


    function randomInt(minimum, maximum) {
        return Math.floor(
            minimum +
            Math.random() * (maximum - minimum + 1)
        );
    }
})();