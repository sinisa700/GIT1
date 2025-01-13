a=1
b=6
print(a+b)

# %%
print("This is a code cell")

# %%
print("This is another code cell")


# %%
import subprocess

# Get the URL of the GitHub repository
def get_git_repo_url():
    result = subprocess.run(['git', 'config', '--get', 'remote.origin.url'], stdout=subprocess.PIPE, text=True)
    print('ovo je print: '+str(result))
    return result.stdout.strip() 
get_git_repo_url()
# %%
