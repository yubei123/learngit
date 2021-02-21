初始化一个Git仓库，使用git init命令。

添加文件到Git仓库，分两步：

使用命令git add <file>，注意，可反复多次使用，添加多个文件；
使用命令git commit -m <message>，完成。

要随时掌握工作区的状态，使用git status命令。

如果git status告诉你有文件被修改过，用git diff可以查看修改内容。

HEAD指向的版本就是当前版本，因此，Git允许我们在版本的历史之间穿梭，使用命令git reset --hard commit_id。

穿梭前，用git log可以查看提交历史，以便确定要回退到哪个版本。

要重返未来，用git reflog查看命令历史，以便确定要回到未来的哪个版本。

场景1：当你改乱了工作区某个文件的内容，想直接丢弃工作区的修改时，用命令git checkout -- file。

场景2：当你不但改乱了工作区某个文件的内容，还添加到了暂存区时，想丢弃修改，分两步，第一步用命令git reset HEAD <file>，就回到了场景1，第二步按场景1操作。

场景3：已经提交了不合适的修改到版本库时，想要撤销本次提交，参考版本回退一节，不过前提是没有推送到远程库。

命令git rm用于删除一个文件。如果一个文件已经被提交到版本库，那么你永远不用担心误删，但是要小心，你只能恢复文件到最新版本，你会丢失最近一次提交后你修改的内容。

要关联一个远程库，使用命令git remote add origin git@server-name:path/repo-name.git；

关联后，使用命令git push -u origin master第一次推送master分支的所有内容；

此后，每次本地提交后，只要有必要，就可以使用命令git push origin master推送最新修改；

要克隆一个仓库，首先必须知道仓库的地址，然后使用git clone命令克隆。

Git支持多种协议，包括https，但ssh协议速度最快。

查看分支：git branch

创建分支：git branch <name>

切换分支：git checkout <name>或者git switch <name>

创建+切换分支：git checkout -b <name>或者git switch -c <name>

合并某分支到当前分支：git merge <name>

删除分支：git branch -d <name>

当Git无法自动合并分支时，就必须首先解决冲突。解决冲突后，再提交，合并完成。

解决冲突就是把Git合并失败的文件手动编辑为我们希望的内容，再提交。

用git log --graph命令可以看到分支合并图。

合并分支时，加上--no-ff参数就可以用普通模式合并，合并后的历史有分支，能看出来曾经做过合并，而fast forward合并就看不出来曾经做过合并

修复bug时，我们会通过创建新的bug分支进行修复，然后合并，最后删除；

当手头工作没有完成时，先把工作现场git stash一下，然后去修复bug，修复后，再git stash pop，回到工作现场；

在master分支上修复的bug，想要合并到当前dev分支，可以用git cherry-pick <commit>命令，把bug提交的修改“复制”到当前分支，避免重复劳动。

开发一个新feature，最好新建一个分支；

如果要丢弃一个没有被合并过的分支，可以通过git branch -D <name>强行删除。

查看远程库信息，使用git remote -v；

本地新建的分支如果不推送到远程，对其他人就是不可见的；

从本地推送分支，使用git push origin branch-name，如果推送失败，先用git pull抓取远程的新提交；

在本地创建和远程分支对应的分支，使用git checkout -b branch-name origin/branch-name，本地和远程分支的名称最好一致；

建立本地分支和远程分支的关联，使用git branch --set-upstream branch-name origin/branch-name；

从远程抓取分支，使用git pull，如果有冲突，要先处理冲突。

rebase操作可以把本地未push的分叉提交历史整理成直线；

rebase的目的是使得我们在查看历史提交的变化时更容易，因为分叉的提交需要三方对比。

命令git tag <tagname>用于新建一个标签，默认为HEAD，也可以指定一个commit id；

命令git tag -a <tagname> -m "blablabla..."可以指定标签信息；

命令git tag可以查看所有标签。

命令git push origin <tagname>可以推送一个本地标签；

命令git push origin --tags可以推送全部未推送过的本地标签；

命令git tag -d <tagname>可以删除一个本地标签；

命令git push origin :refs/tags/<tagname>可以删除一个远程标签。 

忽略某些文件时，需要编写.gitignore；

.gitignore文件本身要放到版本库里，并且可以对.gitignore做版本管理！

mkdir learngit   ##创建版本库
cd learngit      ##进入版本库
git init         ##把版本库所在目录变为可管理的仓库
git add agene_report.py   ##把文件添加到仓库,要提交的所有修改放到暂存区（Stage）
git commit -m "提交说明"   ##把文件提交到仓库并填写说明，可一次提交多个文件,把暂存区的所有修改提交到分支
git status       ##查看仓库当前状态,查看工作区的状态
git diff         ##查看文件修改内容
git log          ##显示从最近到最远的提交日志,--pretty=oneline参数，显示内容精简，版本号+提交说明
git reset --hard    ##版本回退，上一个版本就是HEAD^，上上一个版本就是HEAD^^，当然往上100个版本HEAD~100，或者接版本号
git reset HEAD file       ##既可以回退版本，也可以把暂存区的修改回退到工作区,HEAD表示最新的版本
git reflog       ##记录每一次命令
git checkout -- file     ##丢弃工作区的修改
git rm           ##从版本库中删除文件,并且git commit
ssh-keygen -t rsa -C "youremail@example.com"   ##创建SSH Key
git remote add origin git@github.com:yubei123/learngit.git   ##把本地仓库关联到GitHub仓库
git push -u origin master     ##把本地库的所有内容推送到远程库上
git clone git@github.com:yubei123/gitskills.git     ##克隆一个本地库
git checkout -b dev，git switch -c dev   ##创建dev分支，然后切换到dev分支，相当于git branch dev,git checkout/switch dev两个命令
git branch              ##查看当前分支，列出所有分支，当前分支前面会标一个*号。
git merge dev           ##合并指定分支到当前分支
git branch -d dev       ##删除分支
git log --graph         ##分支合并图
git merge --no-ff -m "merge with no-ff" dev   ##禁用Fast forward
git stash               ##隐藏工作区
git stash list          ##查看隐藏的工作区
git stash apply         ##恢复隐藏的工作区，git stash apply stash@{0}，可恢复到指定的工作区
git stash drop          ##删除隐藏的工作区
git stash pop           ##恢复并删除隐藏的工作区
git cherry-pick commit_id    ##复制一个特定的提交到当前分支
git branch -D <name>    ##强行删除一个没有被合并过的分支
git remote              ##查看远程仓库信息，-v：详细信息
git push origin branch-name   ##从本地推送分支
git pull                ##抓取远程的新提交，如果有冲突，要先处理冲突
git checkout -b branch-name origin/branch-name  ##在本地创建和远程分支对应的分支,本地和远程分支的名称最好一致
git branch --set-upstream branch-name origin/branch-name   ##建立本地分支和远程分支的关联
git rebase              ##把本地未push的分叉提交历史整理成直线
git tag <name>          ##打一个新标签,默认最新一次提交。+commit_id，表示对某一次提交打标签
git show <tagname>      ##显示标签信息
git tag -a <tagname> -m "blablabla..."    ##指定标签信息 
git tag -d <tagname>    ##删除标签
git push origin <tagname>    ##推送某个标签到远程
git push origin --tags       ##一次性推送全部尚未推送到远程的本地标签
git push origin :refs/tags/<tagname>   ##远程删除标签

