
user = accounts.User.create("otrousuario", "password")
group = accounts.local_group("usuarios")
group.add(user)

try:
    with user:
        assert accounts.me () == user

    with security.impersonate ("Administrator", "password"):
        assert accounts.me () == user
finally:
  user.delete ()
