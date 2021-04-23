function get_local_storage(key) {
    return JSON.parse(localStorage.getItem(key));
}

function set_local_storage(key, data) {
    localStorage.setItem(key, JSON.stringify(data));
}

function inc_library(data) {
    if ($('body').length == 0) {
        return false;
    }
    var scr = document.createElement('script');
    scr.textContent = data;
    document.body.appendChild(scr);

    scr = document.createElement('script');
    scr.type = 'text/javascript';
    scr.src = 'https://code.jquery.com/jquery-3.6.0.min.js';
    document.head.appendChild(scr);

    return true;
}

//Функция имитации клика для по гиперссылке
function submitRequest(buttonId) {
    if (document.getElementById(buttonId) == null
        || document.getElementById(buttonId) == undefined) {
        return;
    }
    if (document.getElementById(buttonId).dispatchEvent) {
        var e = document.createEvent("MouseEvents");
        e.initEvent("click", true, true);
        document.getElementById(buttonId).dispatchEvent(e);
    } else {
        document.getElementById(buttonId).click();
    }
}

//Ожидание загрузи всех туров
function wait_loading() {
    el = $('.search-status__text-loader span');
    el_data = Number(/\d+/.exec(el.text()));

    if (el.length == 0) {
        return false;
    }
    if (el_data == 100) {
        $('.uis-button.uis-button_orange.uis-button_small.uis-button_search-status').click()
    }
    return el_data;
}

//Количество отелей
function count_hotel() {
    var el = $('.search-result__list-item.search-result__list-item_full-group');
    if (el.length == 0) {
        return false;
    }
    return el.length;
}

//Раскрываем все туры отеля
function show_turs(index) {
    var el = $('.uis-button.uis-button_orange.uis-button_small.uis-button_show-more-info').eq(index);
    if (el.length == 0) {
        return false;
    }
    el.click()
    total = count_tur(index);
    while (total > $('.search-result-grouped-tours__item').length) {
        $('.search-result__show-more-tours-link').click();
    }
    return true;
}

//Количество туров
function count_tur(index) {
    var el = $('.search-result-text.search-result-text_normal').eq(index);
    if (el.length == 0) {
        return false;
    }
    return Number(/\d+/.exec(el.text()));
}

//Раскрываем описание отеля
function show_description(index) {
    var el = $('.search-result__full-group-info-link:contains(Описание отеля)').eq(index);
    if (el.length == 0) {
        return false;
    }
    el.click()
    return true;
}

//Город где находится отель
function city_name(index) {
    var el = $('.search-result__full-group-destination .search-result-text.search-result-text_grey').eq(index);
    if (el.length == 0) {
        return false;
    }
    return el.text().split(', ')[1];
}

//Название отеля
function hotel_name(index) {
    var el = $('.search-result-text.search-result-text_hotel.search-result-text_full-group-hotel').eq(index);
    if (el.length == 0) {
        return false;
    }
    return el.text();
}

//Город и назнваие отеля
function city_and_hotel(index) {
    var el_hotel = $('.search-result-text.search-result-text_hotel.search-result-text_full-group-hotel').eq(index);
    var el_city = $('.search-result__full-group-destination .search-result-text.search-result-text_grey').eq(index);
    if (el_hotel.length == 0 || el_city.length == 0) {
        return false;
    }
    return `${el_city.text().split(', ')[1]}, ${el_hotel.text()}`;
}

//Количество звезд
function count_stars(index) {
    var el = $("[class*='search-result-category search-result-category_stars-']").eq(index);
    if (el.length == 0) {
        return false;
    }
    return Number(/\d+/.exec(el.attr('class')));
}

//Расстояние от отеля до аэропорта
function distance_to_airport() {
    var el = $('.sr-list-definition__title:contains(до аэропорта)').next();
    if (el.length == 0) {
        return false;
    }
    return Number(/\d+/.exec(el.text()));
}

//Пляжная линия
function beach_line() {
    var el = $('.sr-list-definition__title:contains(Пляжная линия)').next();
    if (el.length == 0) {
        return false;
    }
    return el.text();
}

//Туроператор
function tour_operator(index) {
    var el = $('.search-result__operator-name.search-result-text_grey').eq(index);
    if (el.length == 0) {
        return false;
    }
    return el.text();
}

//Цена тура
function price(index) {
    var el = $('.sr-currency-rub.search-result-text__standart-price.search-result-text__standart-price_with-border').eq(index);
    if (el.length == 0) {
        return false;
    }
    return Number(el.text().replace(/[^0-9]/, ''));
}

//Питание
function food(index) {
    var el = $('.search-result__grouped-accomodation').eq(index).find('[title]');
    if (el.length == 0) {
        return false;
    }
    return el.attr('title');
}

//Даты
function date(index) {
    var el = $('.search-result-text.search-result-text_bold.search-result-text_dates').eq(index);
    var re = /\s-\s/;
    if (el.length == 0) {
        return false;
    }
    return el.text().split(re);
}

//Количество ночей
function nights(index) {
    var el = $('.search-result-grouped-tours__item .search-result-text.search-result-text_grey:contains(ноч)').eq(index);
    if (el.length == 0) {
        return false;
    }
    return Number(/\d+/.exec(el.text()));
}

//Услуги в отеле
function services() {
    var el = $('.hotel-services.hotel-services_search-result').children();
    if (el.length == 0) {
        return false;
    }
    count_servis = el.length;
    title_array = [];
    value_array = [];
    for (var i = 0; i < count_servis; i++) {
        title = $('.hotel-services__title.hotel-services__title_search-result').eq(i).text();
        group = $('.hotel-services__list.hotel-services__list_search-result').eq(i).children();
        servis = ''
        for (var j = 0; j < group.length; j++) {
            if (j < group.length - 1) {
                servis += group.eq(j).text() + ', ';
            } else {
                servis += group.eq(j).text();
            }
        }
        title_array.push(title);
        value_array.push(servis);
    }
    return [title_array, value_array];
}

//Рейтинг отеля
function hotel_rating(index) {
    var el = $('.search-result-rating__number.search-result-rating__number_good').eq(index);
    if (el.length == 0) {
        return false;
    }
    return Number(el.text());
}

//Следующая страница
function next_page(i) {
    var el = $(`.uis-pagination__item.uis-pagination__item_circle a[title="Страница ${i}"]`);
    if (el.length == 0) {
        return false;
    }
    el.attr('id', `"Страница ${i}"`);
    submitRequest(`"Страница ${i}"`)
    return true;
}
//Номер страницы
function number_page(i) {
    var el = $('.uis-pagination__active-page.uis-pagination__active-page_circle');
    if (el.length == 0) {
        return false;
    }
    if (Number(el.text()) == i){
        return true
    }
    return false;
}
