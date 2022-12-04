# CleverCST
该程序实现了CST中的遗传算法和连续自动化计算，目前仅在2015版本上进行测试。  
本程序按照如下方式架构：  
* 遗传算法配置读取，子程序，ReadConfig  
* 参数配置读取，子程序，ReadParConfig  
* 遗传算法过程，子程序，DoGA  
	* 初始参数生成，子程序，BeginForPar  
	* 个体估价，函数，CalcValue  
		* 初始化，自带方法  
		* 建立模型，子程序，MakeModel  
		* 设定参数，子程序，SetPar  
		* 运行计算并导出1D结果，子程序，CalcAndGetResult1D  
		* 计算谐振频率，函数，GetFre  
		（补：初始化，自带代码）  
		* 建立模型，子程序，MakeModel   
		* 设定频率，子程序，SetFre  
		* 运行计算并导出3D结果，子程序，CalcAndGetResult2D  
		* 计算估价函数，函数，PG  
	* 遗传算法生成新个体，函数，MakeNewBody  
	* 输出该代概况，子程序，OutputProfile  
  
在不同情形下需要修改的代码（\*表示可能不需要被修改）：  
1.仍为遗传算法：MakeModel，CalcAndGetResult1D*，GetFre*，SetFre*，CalcAndGetResult2D*，PG  
2.同一问题不同算法：ReadConfig，BeginForPar*，MakeNewBody  
  
值得注意的：  
1.遗传过程中前一代的最优个体直接参与  
2.需要提前进行的配置：config.txt，par_config.txt，1D监视器名称常量，2D监视器名称常量，tempdir（临时目录常量）  
  
  
# 更新日志：  
## 3.0版本  
1.加入了按照频率求算S参数的功能  
2.修复了S参数导出时有概率导出不同分量的问题（提前进行设定）  
3.修复了建模有概率失败导致求解出错直接跳出程序的错误（利用On Error错误捕捉）  
4.修复了参数设定改变后模型不会更新的问题（正确做法是先运行参数设定再运行建模过程）  
  
## 4.0版本：  
1.加入按照相位评估适应度的功能，驻波将会得到0.2的惩罚因子  
2.修复输出每代概况时滞后一代的问题  
3.修改了产生新个体的方法，降低陷入局部最优的概率  
  
## 5.0版本：  
1.增加按照远场评估适应度的功能  
2.增加日志文件记录功能，防止断电等意外导致前功尽弃  
