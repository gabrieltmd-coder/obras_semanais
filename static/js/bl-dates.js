/**
 * bl-dates.js — Validador reutilizável de datas de Baseline
 * Regra: SE Real > Fim BL → Forecast obrigatório
 *
 * Uso: marcar o container pai com data-bl-group="<id_único>"
 * Inputs dentro do grupo devem ter as classes:
 *   .bl-inicio  .bl-fim  .bl-real  .bl-forecast
 */

(function () {
    'use strict';

    /** Valida um único grupo BL. Retorna true se válido. */
    function validarGrupo(group) {
        const fimEl  = group.querySelector('.bl-fim');
        const realEl = group.querySelector('.bl-real');
        const fcEl   = group.querySelector('.bl-forecast');
        if (!fimEl || !realEl || !fcEl) return true;

        const fim      = fimEl.value;
        const real     = realEl.value;
        const forecast = fcEl.value;
        const needFC   = real && fim && real > fim;

        const fcField = fcEl.closest('.bl-field') || fcEl.parentElement;
        const msg     = fcField.querySelector('.bl-msg');
        const reqMark = fcField.querySelector('.bl-req-mark');

        fcField.classList.toggle('bl-required', !!needFC);
        if (reqMark) reqMark.style.display = needFC ? 'inline' : 'none';

        if (needFC && !forecast) {
            fcField.classList.add('bl-err');
            if (msg) msg.style.display = 'block';
            fcEl.setCustomValidity('Forecast obrigatório quando Real > Fim BL');
            return false;
        }

        fcField.classList.remove('bl-err');
        if (msg) msg.style.display = 'none';
        fcEl.setCustomValidity('');
        return true;
    }

    /** Valida todos os grupos BL dentro de um container (seletor CSS ou elemento). */
    window.validarTodasBL = function (containerSel) {
        const root = typeof containerSel === 'string'
            ? document.querySelector(containerSel)
            : (containerSel || document);
        let ok = true;
        (root || document).querySelectorAll('[data-bl-group]').forEach(g => {
            if (!validarGrupo(g)) ok = false;
        });
        return ok;
    };

    /** Valida um único grupo pelo seletor do container ou elemento. */
    window.validarGrupoBL = validarGrupo;

    /** Coleta os 4 valores BL de um grupo. */
    window.coletarBL = function (group) {
        return {
            inicio_bl: group.querySelector('.bl-inicio')?.value.trim() || null,
            fim_bl:    group.querySelector('.bl-fim')?.value.trim()    || null,
            real:      group.querySelector('.bl-real')?.value.trim()   || null,
            forecast:  group.querySelector('.bl-forecast')?.value.trim() || null,
        };
    };

    /* Auto-wire: escuta mudanças em qualquer .bl-real ou .bl-fim da página */
    document.addEventListener('input', function (e) {
        if (!e.target.classList.contains('bl-real') &&
            !e.target.classList.contains('bl-fim')) return;
        const group = e.target.closest('[data-bl-group]');
        if (group) validarGrupo(group);
    });

    /* Inicializa o estado visual de todos os grupos ao carregar */
    document.addEventListener('DOMContentLoaded', function () {
        document.querySelectorAll('[data-bl-group]').forEach(validarGrupo);
    });
}());
