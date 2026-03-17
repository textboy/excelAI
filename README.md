# ExcelAI
ExcelAI is a Microsoft Excel add-in that uses machine learning to manipulate your data by AI techniques.

---
## Requirements
```shell
cd ExcelAI
npm install dotenv --save-dev
npm install remarkable --save
```

## Testing
1. Update your OpenRouter API key in .env
2. Test command
```shell
cd ExcelAI
npm start
```
3. Test data
Put below text into column A1
```
predict: 7,14,21,28,35
may be: 17,41,12,82,53
```
4. Select your preferred model
5. Test prompt in ExcelAI
```
extract final 5 figures of the cells, put those 5 figures into column B to F separated by comma.
```
6. Test result
Refer to the screenshot [test result](sample/test_result.png)

## Debug
```shell
cd ExcelAI
npm run dev-server
```

## Troubleshooting
Error: Access Denied
```
Run as Administrator
```

Error: The dev server is not running on port 3000
```shell
netstat -ano | findstr :3000
taskkill /F /PID <pid>
```
or
```shell
taskkill /F /IM node.exe
```

## Production
1. Build the project
```shell
cd ExcelAI
npm run build
```
2. Update ./manifest.xml, from "https://localhost:3000/" to <local path> (e.g. C:\ExcelAI\)
3. Click install.cmd