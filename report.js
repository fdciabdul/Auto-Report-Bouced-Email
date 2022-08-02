
const excel = require('exceljs');
const nodemailer = require("nodemailer");
// create nodejs to connect ssh and send command
const mysql = require("mysql2/promise");

var format = require("date-format");
var CronJob = require('cron').CronJob;
var SSH2Promise = require('ssh2-promise');
var sshconfig = {
    host: 'host',
    port: 22,
    username: 'user',
    password: 'pass'
  }
  const transporter = nodemailer.createTransport({
    host: 'mail.ptkei.co.id',
    port: 25,
    auth: {
        user: 'user',
        pass: 'pass',
    },
    tls: {
        rejectUnauthorized: false
    }
});
  const createConnection = async () => {
    return await mysql.createConnection({
        host: "localhost",
        user: "root",
        password: "",
        database: "db_smtp_bounce_backup"
    });
}

  var ssh = new SSH2Promise(sshconfig);
  
  //Promise
  ssh.connect().then(() => {
    console.log("Connection established") 
  });
async function execSSH(){
    let data = await ssh.exec(`cat /var/log/mail.log |grep "$(date -d '-24 hour' '+%b %e %H').*postfix/smtp.*status=bounced"`);
     data += await ssh.exec(`cat /var/log/mail.log |grep "$(date -d '-23 hour' '+%b %e %H').*postfix/smtp.*status=bounced"`);

     data += await ssh.exec(`cat /var/log/mail.log |grep "$(date -d '-22 hour' '+%b %e %H').*postfix/smtp.*status=bounced"`);
     data += await ssh.exec(`cat /var/log/mail.log |grep "$(date -d '-20 hour' '+%b %e %H').*postfix/smtp.*status=bounced"`);
     data += await ssh.exec(`cat /var/log/mail.log |grep "$(date -d '-21 hour' '+%b %e %H').*postfix/smtp.*status=bounced"`);
     data += await ssh.exec(`cat /var/log/mail.log |grep "$(date -d '-19 hour' '+%b %e %H').*postfix/smtp.*status=bounced"`);
     data += await ssh.exec(`cat /var/log/mail.log |grep "$(date -d '-18 hour' '+%b %e %H').*postfix/smtp.*status=bounced"`);
     data += await ssh.exec(`cat /var/log/mail.log |grep "$(date -d '-17 hour' '+%b %e %H').*postfix/smtp.*status=bounced"`)
     data += await ssh.exec(`cat /var/log/mail.log |grep "$(date -d '-16 hour' '+%b %e %H').*postfix/smtp.*status=bounced"`)
     data += await ssh.exec(`cat /var/log/mail.log |grep "$(date -d '-15 hour' '+%b %e %H').*postfix/smtp.*status=bounced"`)
     data += await ssh.exec(`cat /var/log/mail.log |grep "$(date -d '-14 hour' '+%b %e %H').*postfix/smtp.*status=bounced"`)
     data += await ssh.exec(`cat /var/log/mail.log |grep "$(date -d '-13 hour' '+%b %e %H').*postfix/smtp.*status=bounced"`)
     data += await ssh.exec(`cat /var/log/mail.log |grep "$(date -d '-12 hour' '+%b %e %H').*postfix/smtp.*status=bounced"`)
     data += await ssh.exec(`cat /var/log/mail.log |grep "$(date -d '-11 hour' '+%b %e %H').*postfix/smtp.*status=bounced"`)
     data += await ssh.exec(`cat /var/log/mail.log |grep "$(date -d '-10 hour' '+%b %e %H').*postfix/smtp.*status=bounced"`)
     data += await ssh.exec(`cat /var/log/mail.log |grep "$(date -d '-9 hour' '+%b %e %H').*postfix/smtp.*status=bounced"`)
     data += await ssh.exec(`cat /var/log/mail.log |grep "$(date -d '-8 hour' '+%b %e %H').*postfix/smtp.*status=bounced"`)
     data += await ssh.exec(`cat /var/log/mail.log |grep "$(date -d '-7 hour' '+%b %e %H').*postfix/smtp.*status=bounced"`)
     data += await ssh.exec(`cat /var/log/mail.log |grep "$(date -d '-6 hour' '+%b %e %H').*postfix/smtp.*status=bounced"`)
     data += await ssh.exec(`cat /var/log/mail.log |grep "$(date -d '-5 hour' '+%b %e %H').*postfix/smtp.*status=bounced"`)
     data += await ssh.exec(`cat /var/log/mail.log |grep "$(date -d '-4 hour' '+%b %e %H').*postfix/smtp.*status=bounced"`)
     data += await ssh.exec(`cat /var/log/mail.log |grep "$(date -d '-3 hour' '+%b %e %H').*postfix/smtp.*status=bounced"`)
     data += await ssh.exec(`cat /var/log/mail.log |grep "$(date -d '-2 hour' '+%b %e %H').*postfix/smtp.*status=bounced"`)
     data += await ssh.exec(`cat /var/log/mail.log |grep "$(date -d '-1 hour' '+%b %e %H').*postfix/smtp.*status=bounced"`)
    return data;
}

function getFirstGroup(regexp, str) {
  return Array.from(str.matchAll(regexp), m => m[1]);
}
var _escapeString = function (val) {
    val = val.replace(/[\0\n\r\b\t\\'"\x1a]/g, function (s) {
        switch (s) {
            case "\0":
                return "\\0";
            case "\n":
                return "\\n";
            case "\r":
                return "\\r";
            case "\b":
                return "\\b";
            case "\t":
                return "\\t";
            case "\x1a":
                return "\\Z";
            case "'":
                return "''";
            case '"':
                return '""';
            default:
                return "\\" + s;
        }
    });

    return val;
};
function start(){
execSSH().then( async(data) => {
     console.log(data)
   var string = data;
    var mailid = string.match(/postfix\/smtp\[\w+\]:\s(\w+)/g);
    var mailto = string.match(/to=<(.*?)>/g);
    var datetime = string.match(/(\w+\s+\w+\s+\w+:\w+:\w+)/g);
    var host = getFirstGroup(/.*?\(host\s(.*?)\[.*/g,string);
    var hostip = getFirstGroup(/.*?\(host\s.*?\[(.*?)\]\s.*/g,string);
    var reason = getFirstGroup(/.*\ssaid:\s(.*)/g, string)
    var parsestring = string.split("\n");
       const connection = await createConnection();
     let number = [];
     for(var i=0; i < host.length; i++){
 const [insert] = await connection.query(
            `INSERT INTO db_bounced (id, date, mail_id, mail_to, host, host_ip, status, reason) 
            VALUES 
            (NULL, '${datetime[i]}', '${mailid[i].split(": ")[1]}', '${_escapeString(mailto[i].split("<")[1].split(">")[0])}', '${host[i]}', 
            '${hostip[i]}', 'Bounced', '${_escapeString(reason[i])}');`
        );
           
              number.push({
                nomor: i,
                date: datetime[i] ,
                mailid: mailid[i].split(": ")[1] ,
                mailto: mailto[i].split("<")[1].split(">")[0] ,
                host: host[i],
                hostip:hostip[i],
                status: "Bounced",
                reason: reason[i]
            })
      }
     

    let workbook = new excel.Workbook(); //creating workbook
        let worksheet = workbook.addWorksheet('Report Bounce '); //creating worksheet
     // worksheet.columns = [
        //  { header: 'No.', key: 'nomor', width: 10 }]

        // worksheet.addRows(number);
        // //  WorkSheet Header
        worksheet.columns = [
           { header: 'No.', key: 'nomor', width: 5 },
            { header: 'DATE TIME', key: 'date', width: 30},
            { header: 'MAIL ID', key: 'mailid', width: 30 },
            { header: 'MAIL TO', key: 'mailto', width: 30},
            { header: 'HOST', key: 'host', width: 30},
            { header: 'HOST IP', key: 'hostip', width: 30, outlineLevel: 1},
            { header: 'STATUS', key: 'status', width: 30},
            { header: 'REASON', key: 'reason', width: 30},
        ];
     
        // Add Array Rows
        worksheet.addRows(number);
     
        // Write to File
          var date = format("yyyy-MM-dd", new Date());
        const buffer = await workbook.xlsx.writeBuffer();
        const mailOptions = {
            from: 'Abdul <abdul@ptkei.co.id>',
     to: ["abdul@ptkei.co.id","teguhsuhartono@ptkei.co.id","fauzan@ptkei.co.id","narainemo@gmail.com"],
  //  cc: ['narainemo@gmail.com','fauzan@ptkei.co.id', 'catur@ptkei.co.id','narainemo@gmail.com','wiwie@ptkei.co.id'],
   subject: 'Auto Report Bounced from SMTP '+ date,
    html: `Dear Team,
    <br><br>
     Silahkan cek Report email bounced berdasarkan 24 jam terakhir ditanggal
     ${date}
    <br><br>
    Demikian informasinya,<br>
    Salam  `,
            attachments: [
                {
                    filename: 'bounce_.xlsx',
                    content: buffer,
                    contentType:
                        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                },
            ],
        };
        const send = await transporter.sendMail(mailOptions);
        console.log(send);
// console.log(string.match(regex))

})
}
 // const job = new CronJob('0 7 * * *', async function() {

 //    try {
 //        // statements
 //        console.log(" =========== STARTING APPS ===========")
         start();
    // } catch(e) {
    //     // statements
    // }
       
    //  });
    // job.start();
   
