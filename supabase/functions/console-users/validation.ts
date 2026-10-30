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
