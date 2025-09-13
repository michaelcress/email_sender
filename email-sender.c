/*
 * email-sender.c
 *
 * Send an HTML email via SMTP using libcurl + XOAUTH2 (OAuth token).
 * No username/password auth.
 *
 * Build:
 *   gcc email-sender.c -o email-sender -lcurl
 *
 * Example:
 *   ./email-sender \
 *     --from you@contoso.com \
 *     --to recipient@example.com \
 *     --subject "OAuth2 Test" \
 *     --username you@contoso.com \
 *     --file body.html \
 *     --token "$(cat access_token.txt)"
 *
 * Defaults:
 *   Server: smtp.office365.com
 *   Port:   587 (STARTTLS)
 */

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <time.h>
#include <unistd.h>
#include <getopt.h>
#include <curl/curl.h>

#define DEFAULT_SMTP_HOST "smtp.office365.com"
#define DEFAULT_SMTP_PORT 587

struct payload {
    const char *data;
    size_t len;
    size_t pos;
};

static size_t payload_source(void *ptr, size_t size, size_t nmemb, void *userp) {
    struct payload *p = (struct payload *)userp;
    size_t cap = size * nmemb;
    if (p->pos >= p->len) return 0;
    size_t n = p->len - p->pos;
    if (n > cap) n = cap;
    memcpy(ptr, p->data + p->pos, n);
    p->pos += n;
    return n;
}

static void rfc2822_date(char *buf, size_t buflen) {
    time_t now = time(NULL);
    struct tm tm;
#if defined(_WIN32)
    gmtime_s(&tm, &now);
#else
    gmtime_r(&now, &tm);
#endif
    strftime(buf, buflen, "%a, %d %b %Y %H:%M:%S +0000", &tm);
}

static int read_file(const char *path, char **out, size_t *out_len) {
    FILE *f = fopen(path, "rb");
    if (!f) return -1;
    if (fseek(f, 0, SEEK_END) != 0) { fclose(f); return -1; }
    long sz = ftell(f);
    if (sz < 0) { fclose(f); return -1; }
    rewind(f);

    char *buf = (char *)malloc((size_t)sz + 1);
    if (!buf) { fclose(f); return -1; }

    size_t n = fread(buf, 1, (size_t)sz, f);
    fclose(f);
    buf[n] = '\0';
    *out = buf;
    *out_len = n;
    return 0;
}

static char *build_message(const char *from, const char *to, const char *subject,
                           const char *html_body, size_t *out_len) {
    char datebuf[64];
    rfc2822_date(datebuf, sizeof(datebuf));

    const char *hdr_fmt =
        "Date: %s\r\n"
        "From: <%s>\r\n"
        "To: <%s>\r\n"
        "Subject: %s\r\n"
        "MIME-Version: 1.0\r\n"
        "Content-Type: text/html; charset=UTF-8\r\n"
        "\r\n";

    size_t head_len = (size_t)snprintf(NULL, 0, hdr_fmt, datebuf, from, to, subject);
    size_t body_len = strlen(html_body);
    size_t need = head_len + body_len;

    char *msg = (char *)malloc(need + 1);
    if (!msg) return NULL;

    int n = snprintf(msg, need + 1, hdr_fmt, datebuf, from, to, subject);
    memcpy(msg + n, html_body, body_len);
    msg[need] = '\0';
    *out_len = need;
    return msg;
}

static void usage(const char *prog) {
    fprintf(stderr,
        "Usage: %s [options]\n"
        "Options:\n"
        "  -s, --server   SMTP server host (default: %s)\n"
        "  -P, --port     SMTP port (default: %d)\n"
        "  -f, --from     Sender email address (required)\n"
        "  -t, --to       Recipient email address (required)\n"
        "  -j, --subject  Email subject (default: \"No subject\")\n"
        "  -u, --username SMTP username (usually your full UPN)\n"
        "  -T, --token    OAuth2 access token string (required)\n"
        "  -F, --file     HTML body file (required)\n"
        "  -h, --help     Show this help\n",
        prog, DEFAULT_SMTP_HOST, DEFAULT_SMTP_PORT);
}

int main(int argc, char **argv) {
    const char *server_host = DEFAULT_SMTP_HOST;
    int server_port = DEFAULT_SMTP_PORT;
    const char *from = NULL, *to = NULL, *subject = "No subject";
    const char *username = NULL, *token = NULL, *filename = NULL;

    static struct option long_opts[] = {
        {"server",   required_argument, 0, 's'},
        {"port",     required_argument, 0, 'P'},
        {"from",     required_argument, 0, 'f'},
        {"to",       required_argument, 0, 't'},
        {"subject",  required_argument, 0, 'j'},
        {"username", required_argument, 0, 'u'},
        {"token",    required_argument, 0, 'T'},
        {"file",     required_argument, 0, 'F'},
        {"help",     no_argument,       0, 'h'},
        {0,0,0,0}
    };

    int opt;
    while ((opt = getopt_long(argc, argv, "s:P:f:t:j:u:T:F:h", long_opts, NULL)) != -1) {
        switch (opt) {
            case 's': server_host = optarg; break;
            case 'P': server_port = atoi(optarg); break;
            case 'f': from        = optarg; break;
            case 't': to          = optarg; break;
            case 'j': subject     = optarg; break;
            case 'u': username    = optarg; break;
            case 'T': token       = optarg; break;
            case 'F': filename    = optarg; break;
            case 'h':
            default:
                usage(argv[0]);
                return 1;
        }
    }

    if (!from || !to || !username || !token || !filename) {
        usage(argv[0]);
        return 1;
    }

    char *body = NULL;
    size_t body_len = 0;
    if (read_file(filename, &body, &body_len) != 0) {
        fprintf(stderr, "Failed to read %s\n", filename);
        return 1;
    }

    size_t msg_len = 0;
    char *message = build_message(from, to, subject, body, &msg_len);
    free(body);
    if (!message) {
        fprintf(stderr, "Failed to build message.\n");
        return 1;
    }

    struct payload p = { message, msg_len, 0 };

    CURL *curl = curl_easy_init();
    if (!curl) {
        fprintf(stderr, "curl_easy_init() failed\n");
        free(message);
        return 1;
    }

    struct curl_slist *rcpts = NULL;
    char url[256];
    snprintf(url, sizeof(url), "smtp://%s:%d", server_host, server_port);

    curl_easy_setopt(curl, CURLOPT_URL, url);
    curl_easy_setopt(curl, CURLOPT_USE_SSL, (long)CURLUSESSL_ALL);
    curl_easy_setopt(curl, CURLOPT_SSLVERSION, (long)CURL_SSLVERSION_TLSv1_2);
    curl_easy_setopt(curl, CURLOPT_MAIL_FROM, from);

    rcpts = curl_slist_append(rcpts, to);
    curl_easy_setopt(curl, CURLOPT_MAIL_RCPT, rcpts);

    curl_easy_setopt(curl, CURLOPT_READFUNCTION, payload_source);
    curl_easy_setopt(curl, CURLOPT_READDATA, &p);
    curl_easy_setopt(curl, CURLOPT_UPLOAD, 1L);

    curl_easy_setopt(curl, CURLOPT_USERNAME, username);
    curl_easy_setopt(curl, CURLOPT_XOAUTH2_BEARER, token);

    curl_easy_setopt(curl, CURLOPT_VERBOSE, 1L);

    char errbuf[CURL_ERROR_SIZE];
    curl_easy_setopt(curl, CURLOPT_ERRORBUFFER, errbuf);
    errbuf[0] = '\0'; // ensure it’s empty initially

    curl_easy_setopt(curl, CURLOPT_PROXY, "");
    curl_easy_setopt(curl, CURLOPT_NOPROXY, "*");

    curl_easy_setopt(curl, CURLOPT_USE_SSL, (long)CURLUSESSL_ALL);
    curl_easy_setopt(curl, CURLOPT_SSLVERSION, (long)CURL_SSLVERSION_TLSv1_2);



    CURLcode res = curl_easy_perform(curl);
    if (res != CURLE_OK) {
        fprintf(stderr, "Send failed: %s\n", curl_easy_strerror(res));
    } else {
        printf("Message sent successfully via %s:%d\n", server_host, server_port);
    }

    curl_slist_free_all(rcpts);
    curl_easy_cleanup(curl);
    free(message);
    return (int)res;
}
