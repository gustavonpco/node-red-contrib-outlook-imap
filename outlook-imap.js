const Imap = require('imap');
const { simpleParser } = require('mailparser');

module.exports = function (RED) {
    function OutlookIMAPNode(config) {
        RED.nodes.createNode(this, config);
        const node = this;

        const imap = new Imap({
            user: config.email,
            xoauth2: msg.token, // O token OAuth2 já provido
            host: 'outlook.office365.com',
            port: 993,
            tls: true,
            tlsOptions: { rejectUnauthorized: false }
        });

        function openInbox(cb) {
            imap.openBox('INBOX', false, cb);
        }

        imap.once('ready', function () {
            openInbox(function (err, box) {
                if (err) {
                    node.error("Erro ao abrir a caixa de entrada: " + err);
                    return;
                }
                const searchCriteria = ['UNSEEN']; // E-mails não lidos
                const fetchOptions = { bodies: '', markSeen: true };

                imap.search(searchCriteria, function (err, results) {
                    if (err) {
                        node.error("Erro na busca de e-mails: " + err);
                        return;
                    }

                    if (results.length === 0) {
                        node.log("Nenhum e-mail não lido encontrado.");
                        return;
                    }

                    const f = imap.fetch(results, fetchOptions);
                    f.on('message', function (msg, seqno) {
                        let emailBuffer = '';
                        msg.on('body', function (stream) {
                            stream.on('data', function (chunk) {
                                emailBuffer += chunk.toString('utf8');
                            });
                        });

                        msg.once('end', function () {
                            simpleParser(emailBuffer, (err, mail) => {
                                if (!err) {
                                    node.log("E-mail recebido: " + mail.subject);
                                    const emailData = {
                                        from: mail.from.text,
                                        subject: mail.subject,
                                        date: mail.date,
                                        body: mail.text
                                    };
                                    node.send({ payload: emailData, original: mail });
                                } else {
                                    node.error("Erro ao analisar e-mail: " + err);
                                }
                            });
                        });
                    });

                    f.once('error', function (err) {
                        node.error("Erro ao buscar e-mails: " + err);
                    });

                    f.once('end', function () {
                        node.log('Finalizada a busca de e-mails.');
                        imap.end();
                    });
                });
            });
        });

        imap.once('error', function (err) {
            node.error("Erro na conexão IMAP: " + err);
        });

        imap.once('end', function () {
            node.log('Conexão IMAP encerrada.');
        });

        node.on('input', function (msg) {
            imap.connect();
        });
    }

    RED.nodes.registerType("Outlook IMAP", OutlookIMAPNode, {
        credentials: {
            email: { type: "text" }
        }
    });
};
