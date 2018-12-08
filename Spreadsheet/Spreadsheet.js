if ((typeof Nskd) == 'undefined') Nskd = {};
if (!Nskd.Spreadsheet) Nskd.Spreadsheet = {};
if (!Nskd.Spreadsheet.XSheet) Nskd.Spreadsheet.XSheet = {};
Nskd.Spreadsheet.XSheet.Onclick = function (event) {
    event = event || window.event;
    target = event.target || event.srcElement;
    if (target.nodeName == 'TD') {
        if (target.getAttribute('class') == 'expand') {
            var tr = target.parentNode;
            var level = getRowLevel(tr);
            if (getNodeTextValue(target) == '+') {
                setNodeTextValue(target, '-');
                do {
                    tr = tr.nextSibling;
                    var subLevel = getRowLevel(tr);
                    if (subLevel == (level + 1)) tr.style.display = '';
                } while (subLevel > level);
            }
            else {
                setNodeTextValue(target, '+');
                do {
                    tr = tr.nextSibling;
                    var subLevel = getRowLevel(tr);
                    if (subLevel >= (level + 1)) {
                        tr.style.display = 'none';
                        var tds = tr.getElementsByTagName('td');
                        var td = tds[subLevel + 1];
                        if (getNodeTextValue(td) == '-') {
                            setNodeTextValue(td, '+');
                        }
                    }
                } while (subLevel > level);
            }
        }
    }

    function getNodeTextValue(node) {
        return node.innerText || node.textContent;
    }

    function setNodeTextValue(node, text) {
        while (node.hasChildNodes()) node.removeChild(node.lastChild);
        node.appendChild(document.createTextNode(text));
    }

    function getRowLevel(tr) {
        return parseInt(getNodeTextValue(tr.getElementsByTagName('td')[0]));
    }
};
