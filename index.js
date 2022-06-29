const puppeteer = require('puppeteer');
const xlsx = require('node-xlsx');
const officegen = require('officegen');
const fs = require('fs');
const workSheetsFromFile = xlsx.parse(`${__dirname}/工作表.xlsx`);
const getAccountArr = () => {
    const arr = [];
    for (const item of workSheetsFromFile) {
        if (item.name == 'Sheet1') {
            const data = item.data;
            const len = data.length;
            for (let i = 1; i < len; i++) {
                const child = data[i];
                if (child[0]) {
                    arr.push({
                        username: child[1].toString(),
                        name: child[0],
                        password: child[2].toString()
                    });
                }
            }
        }
    }
    return arr;
};
const accountArr = getAccountArr();
console.log(accountArr);
// const accountArr =
//     {
//         username: '420902xxxxxxxx2253',
//         password: '123456',
//         name: '张三'
//     }
// ;
const pageConfig = {
    headless: true,
    args: ['--start-fullscreen'],
    // args: ['--start-maximized'],
    defaultViewport: {
        width: 1920,
        height: 1080
    }
};
const pageUrl = 'http://whsz.e.eceping.net/#/home';
const delay = 100;

const getScreenshot = async (page, clip, path) => {
    await page.screenshot({
        path,
        clip
    });
};

const wait = async time => {
    return new Promise(resolve => {
        setTimeout(() => {
            resolve(true);
        }, time * 1000 || 2000);
    });
};

// 启动
const launchFun = async () => {
    const browser = await puppeteer.launch(pageConfig);
    await startFun(browser);
    await browser.close();
};

launchFun();

// 开始
const startFun = async browser => {
    let i = 0;
    for (; i < accountArr.length; i++) {
        const page = await browser.newPage();
        await page.goto(pageUrl);
        await wait();
        console.log('===================================');
        console.log(`${accountArr[i].name} 登录中...`);
        await loginFun(page, accountArr[i]);

        console.log('跳转练习记录...');
        await goRecord(page);

        await getResult(page, accountArr[i]);
        console.log('===================================');
        console.log();
    }
    console.log('全部处理完毕！');
};

//登录
const loginFun = async (page, accountArr) => {
    //打开弹窗
    await page.waitForSelector('.nav-login');
    const loginBtn = await page.$('.nav-login');
    await loginBtn.click();

    await wait(1);

    const username = await page.$('input[placeholder=请输入用户名]');
    const password = await page.$('input[placeholder=请输入密码]');

    //输入账号密码
    await username.type(accountArr.username, { delay });
    await password.type(accountArr.password, { delay });

    //登录
    const login = await page.$('.el-button--primary');
    await login.click();
};

//跳转练习记录
const goRecord = async page => {
    await page.waitForSelector('#tab-myExercise');
    const record = await page.$('#tab-myExercise');
    await record.click();
};

//查看题目结果
const getResult = async (page, accountArr) => {
    const myExerciseList = await page.$$('.myExerciseList li');
    let i = 0;
    let docx = officegen('docx');
    docx.on('finalize', written => {
        console.log('文档创建成功');
    });
    for (; i < myExerciseList.length; i++) {
        if (i > 1) {
            return;
        }
        const path = `./截图/${accountArr.name}/第${i + 1}套题/`;
        fs.mkdirSync(path, { recursive: true }, err => {});
        const list = await page.$$('.myExerciseList li .el-button--primary');
        const item = await list[i];
        item.click();
        await wait();
        console.log(`开始处理${accountArr.name}的第${i + 1}套题...`);
        await getTopic(page, path, docx);
    }
    let out = fs.createWriteStream(
        `截图/${accountArr.name}/${accountArr.name}.docx`
    );
    docx.generate(out);
    console.log('文档保存成功');
    await page.evaluate(() => {
        return window.localStorage.clear();
    });
};

//截图题目
const getTopic = async (page, path, docx) => {
    const list = await page.$$('.answer-sheet-list li a');
    let i = 0;
    let pObj = docx.createP();
    console.log('   开始截图...');
    for (; i < list.length; i++) {
        await list[i].click();
        await wait(1);
        const examTopic = await page.$('.examTopic');
        const answerKeys = await page.$$('.answer-keys');
        const examTopicBox = await examTopic.boundingBox();
        const questionBox = examTopicBox;
        if (answerKeys.length !== 1) {
            const answerKeysBox = await answerKeys[0].boundingBox();
            questionBox.height =
                examTopicBox.height + answerKeysBox.height + 20;
        }
        await getScreenshot(page, questionBox, `${path}第${i + 1}题.png`);
        console.log(`       第${i + 1}题截图成功`);
        pObj.addImage(`${path}第${i + 1}题.png`, {
            cx: questionBox.width * 0.7,
            cy: questionBox.height * 0.7
        });
        console.log(`       第${i + 1}题写入文档成功`);
    }
    console.log('   题目截图、写入完毕');
    console.log();
    docx.putPageBreak();
    docx.putPageBreak();
    await page.goto('http://whsz.e.eceping.net/#/studyCentre');
    await wait();
    await goRecord(page);
};
