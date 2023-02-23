# Chartzy 📈 
An Excel add-in for automatically generating charts from tables.

Version: 1.0.0 alpha

## Tools
- React 
- Fluent UI
- OfficeJS
- Jest 
- Cypress
- Webpack

----

## Usage 
To test your add-in in Excel on a browser, run the following command in the root directory of your project. When you run this command, the local web server starts. Replace "{url}" with the URL of an Excel document on your OneDrive or a SharePoint library to which you have permissions.

```
npm run start:web -- --document {url}
```

The following are examples.

```
npm run start:web -- --document https://contoso.sharepoint.com/:t:/g/EZGxP7ksiE5DuxvY638G798BpuhwluxCMfF1WZQj3VYhYQ?e=F4QM1R
```

```
npm run start:web -- --document https://1drv.ms/x/s!jkcH7spkM4EGgcZUgqthk4IK3NOypVw?e=Z6G1qp
```

```
npm run start:web -- --document https://contoso-my.sharepoint-df.com/:t:/p/user/EQda453DNTpFnl1bFPhOVR0BwlrzetbXvnaRYii2lDr_oQ?e=RSccmNP
```