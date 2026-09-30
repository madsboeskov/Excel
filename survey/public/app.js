'use strict';

(function () {
  var form = document.getElementById('survey-form');
  var questionsEl = document.getElementById('questions');
  var formError = document.getElementById('form-error');
  var submitBtn = document.getElementById('submit');
  var card = document.getElementById('card');
  var survey = null;

  function el(tag, attrs, children) {
    var node = document.createElement(tag);
    Object.keys(attrs || {}).forEach(function (key) {
      if (key === 'text') node.textContent = attrs[key];
      else if (key === 'html') node.innerHTML = attrs[key];
      else if (attrs[key] != null && attrs[key] !== false) node.setAttribute(key, attrs[key]);
    });
    (children || []).forEach(function (child) { if (child) node.appendChild(child); });
    return node;
  }

  function fieldName(question) { return 'q_' + question.id; }

  function labelBlock(question, wrap) {
    var label = el('span', { class: 'q__label', id: fieldName(question) + '-label' }, [
      document.createTextNode(question.label),
      question.required ? el('span', { class: 'q__req', text: '*', 'aria-hidden': 'true' }) : null,
    ]);
    wrap.appendChild(label);
    if (question.help) wrap.appendChild(el('span', { class: 'q__help', text: question.help }));
  }

  function textCounter(input, max, wrap) {
    var counter = el('div', { class: 'counter' });
    var update = function () { counter.textContent = input.value.length + ' / ' + max + ' tegn'; };
    input.addEventListener('input', update);
    update();
    wrap.appendChild(counter);
  }

  function buildQuestion(question) {
    var wrap = el('div', { class: 'q', 'data-id': question.id });
    var id = fieldName(question);
    var type = question.type;

    if (type === 'consent') {
      var consentInput = el('input', { type: 'checkbox', id: id, name: id });
      var consentChoice = el('label', { class: 'choice', for: id }, [
        consentInput,
        el('span', {}, [
          document.createTextNode(question.label),
          question.required ? el('span', { class: 'q__req', text: '*', 'aria-hidden': 'true' }) : null,
        ]),
      ]);
      consentInput.addEventListener('change', function () {
        consentChoice.classList.toggle('choice--checked', consentInput.checked);
      });
      if (question.help) wrap.appendChild(el('span', { class: 'q__help', text: question.help }));
      wrap.appendChild(consentChoice);
      wrap.appendChild(el('p', { class: 'q__error', id: id + '-error', hidden: 'hidden', role: 'alert' }));
      return wrap;
    }

    labelBlock(question, wrap);

    if (type === 'textarea') {
      var area = el('textarea', {
        id: id, name: id, class: 'q__field', maxlength: question.maxLength,
        placeholder: question.placeholder || '', 'aria-describedby': id + '-error',
      });
      wrap.appendChild(area);
      textCounter(area, question.maxLength, wrap);
    } else if (type === 'select') {
      var select = el('select', { id: id, name: id, class: 'q__field', 'aria-describedby': id + '-error' }, [
        el('option', { value: '', text: 'Vælg …' }),
      ]);
      question.options.forEach(function (option) {
        select.appendChild(el('option', { value: option.value, text: option.label }));
      });
      wrap.appendChild(select);
    } else if (type === 'radio' || type === 'checkbox') {
      var group = el('div', { class: 'choices', role: type === 'radio' ? 'radiogroup' : 'group', 'aria-labelledby': id + '-label' });
      question.options.forEach(function (option, index) {
        var optionId = id + '-' + index;
        var input = el('input', { type: type === 'radio' ? 'radio' : 'checkbox', id: optionId, name: id, value: option.value });
        var choice = el('label', { class: 'choice', for: optionId }, [input, el('span', { text: option.label })]);
        input.addEventListener('change', function () {
          if (type === 'radio') {
            group.querySelectorAll('.choice').forEach(function (node) { node.classList.remove('choice--checked'); });
          }
          choice.classList.toggle('choice--checked', input.checked);
        });
        group.appendChild(choice);
      });
      wrap.appendChild(group);
    } else if (type === 'rating') {
      var scale = el('div', { class: 'rating', role: 'radiogroup', 'aria-labelledby': id + '-label' });
      for (var value = 1; value <= question.max; value += 1) {
        var ratingId = id + '-' + value;
        var ratingInput = el('input', { type: 'radio', id: ratingId, name: id, value: String(value) });
        var ratingChoice = el('label', { class: 'choice', for: ratingId }, [ratingInput, el('span', { text: String(value) })]);
        (function (input, choice) {
          input.addEventListener('change', function () {
            scale.querySelectorAll('.choice').forEach(function (node) { node.classList.remove('choice--checked'); });
            choice.classList.add('choice--checked');
          });
        })(ratingInput, ratingChoice);
        scale.appendChild(ratingChoice);
      }
      wrap.appendChild(scale);
      if (question.minLabel || question.maxLabel) {
        wrap.appendChild(el('div', { class: 'rating__ends' }, [
          el('span', { text: question.minLabel || '' }),
          el('span', { text: question.maxLabel || '' }),
        ]));
      }
    } else {
      var input = el('input', {
        type: type === 'number' ? 'number' : type,
        id: id, name: id, class: 'q__field',
        placeholder: question.placeholder || '',
        maxlength: question.maxLength || false,
        min: question.min != null ? question.min : false,
        max: question.max != null ? question.max : false,
        autocomplete: type === 'email' ? 'email' : (question.id === 'navn' || question.id === 'name' ? 'name' : 'on'),
        'aria-describedby': id + '-error',
      });
      wrap.appendChild(input);
    }

    wrap.appendChild(el('p', { class: 'q__error', id: id + '-error', hidden: 'hidden', role: 'alert' }));
    return wrap;
  }

  function collect() {
    var data = { _hp: document.getElementById('_hp').value };
    survey.questions.forEach(function (question) {
      var name = fieldName(question);
      var nodes = form.querySelectorAll('[name="' + name + '"]');
      if (question.type === 'checkbox') {
        data[question.id] = Array.prototype.filter.call(nodes, function (node) { return node.checked; })
          .map(function (node) { return node.value; });
      } else if (question.type === 'consent') {
        data[question.id] = nodes[0] ? nodes[0].checked : false;
      } else if (question.type === 'radio' || question.type === 'rating') {
        var picked = Array.prototype.find.call(nodes, function (node) { return node.checked; });
        data[question.id] = picked ? picked.value : '';
      } else {
        data[question.id] = nodes[0] ? nodes[0].value : '';
      }
    });
    return data;
  }

  function clearErrors() {
    formError.hidden = true;
    formError.textContent = '';
    questionsEl.querySelectorAll('.q').forEach(function (node) {
      node.classList.remove('q--invalid');
      var message = node.querySelector('.q__error');
      if (message) { message.hidden = true; message.textContent = ''; }
    });
  }

  function showErrors(errors) {
    var first = null;
    Object.keys(errors || {}).forEach(function (id) {
      var node = questionsEl.querySelector('.q[data-id="' + id + '"]');
      if (!node) return;
      node.classList.add('q--invalid');
      var message = node.querySelector('.q__error');
      if (message) { message.textContent = errors[id]; message.hidden = false; }
      if (!first) first = node;
    });
    if (first) {
      first.scrollIntoView({ behavior: 'smooth', block: 'center' });
      var focusable = first.querySelector('input, select, textarea');
      if (focusable) focusable.focus({ preventScroll: true });
    }
  }

  // Samme regler som serveren, så de fleste fejl fanges før netværkskaldet.
  function validateLocally(data) {
    var errors = {};
    survey.questions.forEach(function (question) {
      var value = data[question.id];
      if (question.type === 'checkbox') {
        if (question.required && value.length === 0) errors[question.id] = 'Vælg mindst én mulighed.';
        return;
      }
      if (question.type === 'consent') {
        if (question.required && !value) errors[question.id] = 'Feltet skal være markeret.';
        return;
      }
      if (question.required && String(value).trim() === '') {
        errors[question.id] = question.type === 'radio' || question.type === 'rating' ? 'Vælg en mulighed.' : 'Feltet skal udfyldes.';
        return;
      }
      if (question.type === 'email' && String(value).trim() !== '' && !/^[^\s@]+@[^\s@.]+(\.[^\s@.]+)+$/.test(value.trim())) {
        errors[question.id] = 'Skriv en gyldig e-mailadresse, fx navn@eksempel.dk.';
      }
    });
    return errors;
  }

  function showThankYou(result) {
    form.hidden = true;
    document.getElementById('done-heading').textContent = result.thankYou.heading;
    document.getElementById('done-body').textContent = result.thankYou.body;
    document.getElementById('mail-warning').hidden = result.mailSent !== false;
    document.getElementById('done').hidden = false;
    card.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  form.addEventListener('submit', function (event) {
    event.preventDefault();
    clearErrors();
    var data = collect();

    var localErrors = validateLocally(data);
    if (Object.keys(localErrors).length > 0) {
      showErrors(localErrors);
      return;
    }

    submitBtn.disabled = true;
    fetch('api/submit', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(data),
    })
      .then(function (response) {
        return response.json().then(function (body) { return { status: response.status, body: body }; });
      })
      .then(function (result) {
        if (result.body && result.body.ok) {
          showThankYou(result.body);
          return;
        }
        formError.textContent = (result.body && result.body.message) || 'Skemaet kunne ikke sendes. Prøv igen.';
        formError.hidden = false;
        showErrors(result.body && result.body.errors);
      })
      .catch(function () {
        formError.textContent = 'Der er ingen forbindelse til serveren. Tjek dit netværk og prøv igen.';
        formError.hidden = false;
      })
      .finally(function () { submitBtn.disabled = false; });
  });

  document.getElementById('again').addEventListener('click', function () {
    form.reset();
    form.querySelectorAll('.choice--checked').forEach(function (node) { node.classList.remove('choice--checked'); });
    form.querySelectorAll('.counter').forEach(function (node) {
      var input = node.parentNode.querySelector('textarea');
      if (input) node.textContent = '0 / ' + input.getAttribute('maxlength') + ' tegn';
    });
    clearErrors();
    document.getElementById('done').hidden = true;
    form.hidden = false;
  });

  fetch('api/survey')
    .then(function (response) { return response.json(); })
    .then(function (data) {
      survey = data;
      document.title = data.title;
      document.getElementById('survey-title').textContent = data.title;
      document.getElementById('survey-intro').textContent = data.intro || '';
      document.getElementById('privacy').textContent = data.privacyNote || '';
      document.querySelector('.submit__label').textContent = data.submitLabel || 'Send';
      data.questions.forEach(function (question) { questionsEl.appendChild(buildQuestion(question)); });
      form.hidden = false;
      card.setAttribute('aria-busy', 'false');
    })
    .catch(function () {
      document.getElementById('survey-title').textContent = 'Spørgeskemaet kunne ikke hentes';
      document.getElementById('survey-intro').textContent = 'Prøv at genindlæse siden. Er serveren startet?';
      card.setAttribute('aria-busy', 'false');
    });
})();
