
function getresults(keyword) {
  let csf = $("input[name=csrfmiddlewaretoken]").val();
  var data = {
    'keyword': keyword,
  };
  if(keyword != ""){
  $.ajax({
    headers: { "X-CSRFToken": csf },
    mode: "same-origin", // Do not send CSRF token to another domain.
    beforeSend: function () {
      $(".preloader").css("visibility", "visible");
    },
    url: "/getresults",
    type: "POST",
    data: data,
    dataType: "html",
    success: function (data) {
      $('.display-results').html("");
      $('.display-results').append(data);
  
    },
    complete: function () {
      $(".preloader").css("visibility", "hidden");
    },
  });
}
};




