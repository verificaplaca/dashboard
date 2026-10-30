export function validateNewUser(value: unknown) {
 if(!value||typeof value!=='object'||Array.isArray(value)) throw new Error('Formulário inválido.');
 const data=value as Record<string,unknown>;
 if(Object.keys(data).some(key=>!['name','email','password','role'].includes(key))) throw new Error('Campos inválidos.');
 if(typeof data.name!=='string'||!data.name.trim()||data.name.trim().length>120) throw new Error('Informe um nome com até 120 caracteres.');
 if(typeof data.email!=='string'||data.email.trim().length>254||!/^\S+@\S+\.\S+$/.test(data.email.trim())) throw new Error('Informe um email válido.');
 if(typeof data.password!=='string'||data.password.length<12||data.password.length>128) throw new Error('Use uma senha entre 12 e 128 caracteres.');
 if(data.role!=='manager'&&data.role!=='administrator') throw new Error('Escolha gerente ou administrador.');
 return {name:data.name.trim(),email:data.email.trim().toLowerCase(),password:data.password,role:data.role};
}
export function validateUserChange(value:unknown){
 if(!value||typeof value!=='object'||Array.isArray(value))throw new Error('Formulário inválido.');
 const data=value as Record<string,unknown>,reset=data.action==='reset_password';
 if(data.action!=='update'&&!reset)throw new Error('Ação inválida.');
 const allowed=reset?['action','id','version','password']:['action','id','version','name','email','role','active'];
 if(Object.keys(data).some(key=>!allowed.includes(key)))throw new Error('Campos inválidos.');
 if(typeof data.id!=='string'||! /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(data.id)||!Number.isInteger(data.version)||Number(data.version)<1)throw new Error('Atualize a lista de usuários.');
 const base={id:data.id,version:Number(data.version)};
 if(reset){
  if(typeof data.password!=='string'||data.password.length<12||data.password.length>128)throw new Error('Use uma senha entre 12 e 128 caracteres.');
  return {...base,action:'reset_password' as const,password:data.password};
 }
 const profile=validateNewUser({name:data.name,email:data.email,role:data.role,password:'validation only 123'});
 if(typeof data.active!=='boolean')throw new Error('Situação inválida.');
 return {...base,action:'update' as const,name:profile.name,email:profile.email,role:profile.role,active:data.active};
}
