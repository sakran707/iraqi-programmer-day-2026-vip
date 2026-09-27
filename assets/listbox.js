/**
 * قائمة منسدلة مخصصة تحل محل <select> الأصلي — لأن قائمة المتصفح/النظام
 * الأصلية (خصوصاً بالوضع الداكن) تظهر بتنسيق منفصل تماماً عن هوية الموقع
 * ولا يمكن التحكم بشكلها عبر CSS.
 */
function createListbox(options, selected, onChange) {
  var wrap = document.createElement('div');
  wrap.className = 'listbox';
  wrap.dataset.value = selected;

  var btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'listbox-btn';
  btn.textContent = selected;

  var menu = document.createElement('div');
  menu.className = 'listbox-menu hidden';

  options.forEach(function (opt) {
    var item = document.createElement('div');
    item.className = 'listbox-option' + (opt === selected ? ' selected' : '');
    item.textContent = opt;
    item.addEventListener('click', function (e) {
      e.stopPropagation();
      btn.textContent = opt;
      wrap.dataset.value = opt;
      menu.classList.add('hidden');
      Array.prototype.forEach.call(menu.querySelectorAll('.listbox-option'), function (o) {
        o.classList.toggle('selected', o === item);
      });
      if (onChange) onChange(opt);
    });
    menu.appendChild(item);
  });

  btn.addEventListener('click', function (e) {
    e.stopPropagation();
    var wasHidden = menu.classList.contains('hidden');
    document.querySelectorAll('.listbox-menu').forEach(function (m) { m.classList.add('hidden'); });
    if (wasHidden) menu.classList.remove('hidden');
  });

  wrap.appendChild(btn);
  wrap.appendChild(menu);
  return wrap;
}

document.addEventListener('click', function () {
  document.querySelectorAll('.listbox-menu').forEach(function (m) { m.classList.add('hidden'); });
});
