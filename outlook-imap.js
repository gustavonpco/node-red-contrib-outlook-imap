const Imap = require('imap-simple');
const { simpleParser } = require('mailparser');

module.exports = function(RED) {
    function OutlookIMAPNode(config) {
        RED.nodes.createNode(this, config);
        const node = this;

        const imapConfig = {
            imap: {
                user: config.email, 
                xoauth2: msg.token, // O Token OAuth2 já provido
                host: 'outlook.office365.com',
                port: 993,
                tls: true,
                tlsOptions: { rejectUnauthorized: false },
                authTimeout: 30000
            }
        };

        node.on('input', function(msg) {
            Imap.connect(imapConfig).then((connection) => {
                return connection.openBox('INBOX').then(() => {
                    const searchCriteria = ['UNSEEN'];
                    const fetchOptions = { bodies: ['HEADER', 'TEXT'], markSeen: true };

                    return connection.search(searchCriteria, fetchOptions).then((messages) => {
                        messages.forEach((item) => {
                            const all = item.parts.find(part => part.which === 'TEXT');
                            const id = item.attributes.uid;
                            const idHeader = "Imap-Id: " + id + "\r\n";
                            simpleParser(idHeader + all.body, (err, mail) => {
                                if (!err) {
                                    msg.payload = {
                                        from: mail.from.text,
                                        subject: mail.subject,
                                        date: mail.date,
                                        body: mail.text
                                    };
                                    node.send(msg);
                                } else {
                                    node.error("Error parsing mail: " + err);
                                }
                            });
                        });
                    });
                });
            }).catch(err => {
                node.error("Error connecting to IMAP: " + err);
            });
        });
    }
    
    RED.nodes.registerType("Outlook IMAP", OutlookIMAPNode, {
        credentials: {
            email: { type: "text" }
        }
    });
}
