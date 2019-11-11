export function showMessage(text) {
    $('.welcome-body').hide();
    $('#message-area').show();
    $('#message-area').text(text);
}