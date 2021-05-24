# zsh下(需要借助oh-my-zsh的ZSH变量来判断是zsh还是bash)
# noglob cscript ./morph.vbs ./assets/test*/*.txt 能够不展开glob
# 交给 cscript ./morph.vbs ./assets/test*/*.txt 内部判断
# 封装方法简单
# alias morph="noglob cscript `realpath ./morph.vbs`"

# bash下
# set -f; cscript ./morph.vbs ./assets/test*/*.txt; set +f 能够不展开blob
# 交给 cscript ./morph.vbs ./assets/test*/*.txt 内部判断
# 封装方法稍微复杂
# reset_expansion(){ CMD="$1";shift;$CMD "$@";set +f;}
# alias morph="set -f; reset_expansion cscript `realpath ./morph.vbs`"

if [[ -n $ZSH ]]; then
    echo "ZSH"
    alias morph="noglob cscript `realpath ./morph.vbs`"
    which morph
else
    echo "BASH"
    reset_expansion(){ CMD="$1";shift;$CMD "$@";set +f;}
    alias morph="set -f; reset_expansion cscript `realpath ./morph.vbs`"
    which morph
fi