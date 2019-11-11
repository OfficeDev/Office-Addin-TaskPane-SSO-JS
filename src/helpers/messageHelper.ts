export function showMessage(text: string): void {
    $('.welcome-body').hide();
    $('#message-area').show();
    $('#message-area').text(text);
}