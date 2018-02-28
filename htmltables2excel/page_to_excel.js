/**
 * Created by chuck on 10/7/15.
 *
 * Creates an excel file by uploading the tables on the page to the server. On the server, the tables are converted
 * to an excel file, then the page calls that link and the file is downloaded.
 *
 * Requires jquery
 *
 * params:
 *
 *  excludes: a list of table IDs to exclude
 *  csrf_token: from Django, you can get it as {{ csrf_token }}
 *  page_to_excel_url: url to call to create excel. If you leave it undefined, then it will call the page its on
 *
 *
 */

make_page_to_excel_func = function(excludes, csrf_token, page_to_excel_url){
    var include;
    var the_tables = [];

    if (page_to_excel_url === undefined){
        page_to_excel_url = window.location;
    }

    // Make a list of tables. Grab the content on demand so that it matches the sorting applied by the user.
    $('table').each(function(i1, a_table){
        include = true;
        $.each(excludes, function(i2, table_id){
            if(table_id === a_table.id){
                include = false;
            }
        });
        if(include) {
            the_tables.push(a_table);
        }
    });

    return function(){
        var table_html = [];

        // Get current table content
        $.each(the_tables, function(i, v){table_html.push($(v).prop('outerHTML'))});

        $.ajax({
            method: "POST",
            url: page_to_excel_url,
            data: {'tables': JSON.stringify(table_html), 'to_excel': true, csrfmiddlewaretoken: csrf_token}
        }).done(function (content) {
            if(content.success){
                /** @namespace content.file_url */
                window.location.href = content.file_url;
            } else {
                alert( "Error creating file");
            }
        });
    }
};
